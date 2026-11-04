import {
  executeRibbonFillAction,
  type FillFormattingMode,
  type FillOptions,
  fillOutputRange,
  fillRange,
} from '../commands/fill.js';
import type { History } from '../commands/history.js';
import { deferForAllowedRangeUnlock } from '../commands/protection.js';
import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import {
  createFloatingOptionsButton,
  createFloatingOptionsMenuItem,
} from './floating-options-menu.js';
import { createObjectMenu, type ObjectMenuEntry } from './object-menu.js';
import { overlayPortalFor } from './overlay-portal.js';
import { clampPanelToViewport } from './overlay-position.js';

type AutoFillMode =
  | 'copy'
  | 'series'
  | 'formattingOnly'
  | 'withoutFormatting'
  | 'days'
  | 'weekdays'
  | 'months'
  | 'years';

/** Chip offset from the last filled cell's bottom-right corner, and the gap between chip and menu. */
const CHIP_OFFSET_X = 1.5;
const CHIP_OFFSET_Y = 2.7;
const MENU_GAP = 5;
const CHIP_FALLBACK_SIZE = 22;

interface AutoFillMenuDetail {
  src: Range;
  dest: Range;
  clientX: number;
  clientY: number;
  anchorX: number;
  anchorY: number;
}

interface AutoFillOptionsDetail {
  src: Range;
  dest: Range;
  mode: AutoFillMode;
  clientX: number;
  clientY: number;
}

export interface AutoFillOptionsHandle {
  detach(): void;
  setStrings(next: Strings): void;
}

export interface AutoFillOptionsDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings: Strings;
  history?: History | null;
  onAfterCommit?: () => void;
}

export function attachAutoFillOptions(deps: AutoFillOptionsDeps): AutoFillOptionsHandle {
  const { host, store, wb } = deps;
  const history = deps.history ?? null;
  let strings = deps.strings;
  let detail: AutoFillOptionsDetail | null = null;
  let menuOpen = false;

  const button = createFloatingOptionsButton({ className: 'fc-autofill-options__button' });

  const menu = document.createElement('div');
  menu.className = 'fc-autofill-options__menu';
  menu.setAttribute('role', 'menu');
  menu.style.display = 'none';

  const createAutoFillItem = (mode: AutoFillMode): HTMLButtonElement =>
    createFloatingOptionsMenuItem({
      className: 'fc-autofill-options__item',
      mode,
      role: 'menuitemradio',
    });

  const copyItem = createAutoFillItem('copy');
  const seriesItem = createAutoFillItem('series');
  const formattingOnlyItem = createAutoFillItem('formattingOnly');
  const withoutFormattingItem = createAutoFillItem('withoutFormatting');
  const dayItems = (['days', 'weekdays', 'months', 'years'] as const).map(createAutoFillItem);

  menu.append(copyItem, seriesItem, formattingOnlyItem, withoutFormattingItem, ...dayItems);
  overlayPortalFor(host).append(button, menu);

  const applyLabels = (): void => {
    const t = strings.autoFillOptions;
    button.title = t.title;
    button.setAttribute('aria-label', t.title);
    menu.setAttribute('aria-label', t.title);
    copyItem.textContent = t.copyCells;
    seriesItem.textContent = t.fillSeries;
    formattingOnlyItem.textContent = t.fillFormattingOnly;
    withoutFormattingItem.textContent = t.fillWithoutFormatting;
    const labels = [t.fillDays, t.fillWeekdays, t.fillMonths, t.fillYears];
    dayItems.forEach((item, index) => {
      item.textContent = labels[index] ?? '';
    });
  };

  const isDateFillCandidate = (): boolean => {
    if (!detail) return false;
    const state = store.getState();
    const { src } = detail;
    const isDate = (key: string): boolean => {
      const fmt = state.format.formats.get(key)?.numFmt;
      return (
        (fmt?.kind === 'date' || fmt?.kind === 'datetime') &&
        state.data.cells.get(key)?.value.kind === 'number'
      );
    };
    // A small source is cheaper to walk cell by cell than the whole format map.
    const area = (src.r1 - src.r0 + 1) * (src.c1 - src.c0 + 1);
    if (area < state.format.formats.size) {
      for (let row = src.r0; row <= src.r1; row += 1) {
        for (let col = src.c0; col <= src.c1; col += 1) {
          if (isDate(addrKey({ sheet: src.sheet, row, col }))) return true;
        }
      }
      return false;
    }
    for (const key of state.format.formats.keys()) {
      if (!isDate(key)) continue;
      const addr = parseAddrKey(key);
      if (!addr || !rangeContainsAddr(detail.src, addr)) continue;
      return true;
    }
    return false;
  };

  const updateDateItems = (): void => {
    const display = isDateFillCandidate() ? '' : 'none';
    for (const item of dayItems) item.style.display = display;
  };

  /** `x`/`y` is the bottom-right corner of the last filled cell; the chip sits just outside it. */
  const position = (x: number, y: number): void => {
    const buttonPos = clampPanelToViewport(button, x + CHIP_OFFSET_X, y + CHIP_OFFSET_Y, {
      fallbackWidth: CHIP_FALLBACK_SIZE,
      fallbackHeight: CHIP_FALLBACK_SIZE,
    });
    button.style.left = `${buttonPos.x}px`;
    button.style.top = `${buttonPos.y}px`;
    const menuPos = clampPanelToViewport(
      menu,
      buttonPos.x,
      buttonPos.y + (button.offsetHeight || CHIP_FALLBACK_SIZE) + MENU_GAP,
      {
        fallbackWidth: 180,
        fallbackHeight: 80,
      },
    );
    menu.style.left = `${menuPos.x}px`;
    menu.style.top = `${menuPos.y}px`;
  };

  const updateChecked = (): void => {
    const mode = detail?.mode ?? 'series';
    copyItem.setAttribute('aria-checked', mode === 'copy' ? 'true' : 'false');
    seriesItem.setAttribute('aria-checked', mode === 'series' ? 'true' : 'false');
    formattingOnlyItem.setAttribute('aria-checked', mode === 'formattingOnly' ? 'true' : 'false');
    withoutFormattingItem.setAttribute(
      'aria-checked',
      mode === 'withoutFormatting' ? 'true' : 'false',
    );
    for (const item of dayItems) {
      item.setAttribute('aria-checked', item.dataset.fcMode === mode ? 'true' : 'false');
    }
  };

  const setMenuOpen = (open: boolean): void => {
    menuOpen = open;
    button.setAttribute('aria-expanded', open ? 'true' : 'false');
    menu.style.display = open ? 'block' : 'none';
    if (open) {
      updateChecked();
      const first = detail?.mode === 'copy' ? copyItem : seriesItem;
      first.focus();
    }
  };

  const hide = (): void => {
    detail = null;
    button.style.display = 'none';
    setMenuOpen(false);
  };

  const formattingModeFor = (mode: AutoFillMode): FillFormattingMode =>
    mode === 'formattingOnly' ? 'only' : mode === 'withoutFormatting' ? 'without' : 'with';

  const dateUnitFor = (mode: AutoFillMode): 'days' | 'weekdays' | 'months' | 'years' | undefined =>
    mode === 'days' || mode === 'weekdays' || mode === 'months' || mode === 'years'
      ? mode
      : undefined;

  /** Run one fill from `src` into `dest` as a single undo step and select the result. */
  const runFill = (src: Range, dest: Range, fillOpts: FillOptions, retry: () => void): boolean => {
    if (deferForAllowedRangeUnlock(store, [fillOutputRange(src, dest)], retry)) return false;
    if (history) history.begin();
    let wrote = false;
    try {
      wrote = fillRange(store.getState(), wb, src, dest, { ...fillOpts, store });
    } finally {
      if (history) history.end();
    }
    if (wrote) {
      deps.onAfterCommit?.();
      mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
      mutators.extendRangeTo(store, { sheet: dest.sheet, row: dest.r1, col: dest.c1 });
    }
    return wrote;
  };

  const reapply = (mode: AutoFillMode): void => {
    const current = detail;
    if (!current) return;
    const wrote = runFill(
      current.src,
      current.dest,
      {
        copyOnly: mode === 'copy',
        series: mode === 'series',
        formatting: formattingModeFor(mode),
        dateUnit: dateUnitFor(mode),
      },
      () => reapply(mode),
    );
    if (wrote) {
      current.mode = mode;
      updateChecked();
    }
    setMenuOpen(false);
    button.focus();
  };

  const dragMenu = createObjectMenu(host);

  /** What the source cells allow, which decides the enabled rows of the drag menu. */
  const sourceKinds = (src: Range): { number: boolean; date: boolean; text: boolean } => {
    const state = store.getState();
    const kinds = { number: false, date: false, text: false };
    for (let row = src.r0; row <= src.r1; row += 1) {
      for (let col = src.c0; col <= src.c1; col += 1) {
        const key = addrKey({ sheet: src.sheet, row, col });
        const value = state.data.cells.get(key)?.value;
        if (value?.kind === 'number') {
          const fmt = state.format.formats.get(key)?.numFmt;
          if (fmt?.kind === 'date' || fmt?.kind === 'datetime') kinds.date = true;
          else kinds.number = true;
        } else if (value?.kind === 'text') {
          // Text ending in a number ("Item 1") still extends as a series.
          if (/\d/.test(value.value)) kinds.number = true;
          else kinds.text = true;
        }
      }
    }
    return kinds;
  };

  const onMenuRequest = (e: Event): void => {
    const next = (e as CustomEvent<AutoFillMenuDetail>).detail;
    if (!next) return;
    const { src, dest } = next;
    const t = strings.autoFillOptions;
    const kinds = sourceKinds(src);
    const series = kinds.number || kinds.date;
    const run =
      (fillOpts: FillOptions): (() => void) =>
      () => {
        const again = (): void => {
          runFill(src, dest, fillOpts, again);
        };
        again();
      };
    const entries: ObjectMenuEntry[] = [
      { id: 'copy', label: t.copyCells, run: run({ copyOnly: true }) },
      { id: 'series', label: t.fillSeries, run: run({ series: true }), disabled: !series },
      {
        id: 'formattingOnly',
        label: t.fillFormattingOnly,
        run: run({ copyOnly: true, formatting: 'only' }),
      },
      {
        id: 'withoutFormatting',
        label: t.fillWithoutFormatting,
        run: run({ series: true, formatting: 'without' }),
      },
      { separator: true },
      ...(['days', 'weekdays', 'months', 'years'] as const).map((unit): ObjectMenuEntry => {
        const label = {
          days: t.dragMenu.seriesDays,
          weekdays: t.dragMenu.seriesWeekdays,
          months: t.dragMenu.seriesMonths,
          years: t.dragMenu.seriesYears,
        }[unit];
        return { id: unit, label, run: run({ dateUnit: unit }), disabled: !kinds.date };
      }),
      { separator: true },
      {
        id: 'linear',
        label: t.dragMenu.linear,
        run: run({ series: true }),
        disabled: !kinds.number,
      },
      {
        id: 'growth',
        label: t.dragMenu.growth,
        run: run({ growth: true }),
        disabled: !kinds.number,
      },
      {
        id: 'flashFill',
        label: t.dragMenu.flashFill,
        disabled: history === null,
        run: () => {
          if (!history) return;
          mutators.setActive(store, { sheet: dest.sheet, row: dest.r0, col: dest.c0 });
          mutators.extendRangeTo(store, { sheet: dest.sheet, row: dest.r1, col: dest.c1 });
          if (executeRibbonFillAction({ store, workbook: wb, history, action: 'flash' })) {
            deps.onAfterCommit?.();
          }
        },
      },
      // The series dialog has no counterpart in this build yet.
      { id: 'seriesDialog', label: t.dragMenu.seriesDialog, run: () => undefined, disabled: true },
    ];
    hide();
    dragMenu.open(next.clientX + 1, next.clientY - 1, entries, {
      label: t.title,
      restoreFocusTo: host,
    });
  };

  const onOpen = (e: Event): void => {
    const next = (e as CustomEvent<AutoFillOptionsDetail>).detail;
    if (!next) return;
    detail = { ...next, src: { ...next.src }, dest: { ...next.dest } };
    applyLabels();
    position(next.clientX, next.clientY);
    updateDateItems();
    updateChecked();
    button.style.display = 'block';
    setMenuOpen(false);
  };

  const onButtonClick = (e: MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    if (!detail) return;
    setMenuOpen(!menuOpen);
  };

  const onMenuClick = (e: MouseEvent): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLButtonElement>(
      '.fc-autofill-options__item',
    );
    if (!target) return;
    e.preventDefault();
    e.stopPropagation();
    const mode = target.dataset.fcMode as AutoFillMode | undefined;
    if (!mode) return;
    reapply(mode);
  };

  const onDocPointerDown = (e: MouseEvent): void => {
    const target = e.target as Node | null;
    if (target && (button.contains(target) || menu.contains(target))) return;
    hide();
  };

  const onDocKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') hide();
  };

  applyLabels();
  host.addEventListener('fc:autofilloptions', onOpen);
  host.addEventListener('fc:autofillmenu', onMenuRequest);
  button.addEventListener('click', onButtonClick);
  menu.addEventListener('click', onMenuClick);
  document.addEventListener('mousedown', onDocPointerDown, true);
  document.addEventListener('keydown', onDocKey, true);
  window.addEventListener('scroll', hide, true);

  return {
    detach(): void {
      host.removeEventListener('fc:autofilloptions', onOpen);
      host.removeEventListener('fc:autofillmenu', onMenuRequest);
      dragMenu.detach();
      button.removeEventListener('click', onButtonClick);
      menu.removeEventListener('click', onMenuClick);
      document.removeEventListener('mousedown', onDocPointerDown, true);
      document.removeEventListener('keydown', onDocKey, true);
      window.removeEventListener('scroll', hide, true);
      button.remove();
      menu.remove();
    },
    setStrings(next: Strings): void {
      strings = next;
      applyLabels();
    },
  };
}
