import type { FillFormattingMode } from '../commands/fill.js';
import { fillRange } from '../commands/fill.js';
import type { History } from '../commands/history.js';
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

type AutoFillMode =
  | 'copy'
  | 'series'
  | 'formattingOnly'
  | 'withoutFormatting'
  | 'days'
  | 'weekdays'
  | 'months'
  | 'years';

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

const VIEWPORT_PAD = 4;

const clamp = (value: number, min: number, max: number): number =>
  Math.max(min, Math.min(max, value));

export function attachAutoFillOptions(deps: AutoFillOptionsDeps): AutoFillOptionsHandle {
  const { host, store, wb } = deps;
  const history = deps.history ?? null;
  let strings = deps.strings;
  let detail: AutoFillOptionsDetail | null = null;
  let menuOpen = false;

  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'fc-autofill-options__button';
  button.setAttribute('aria-haspopup', 'menu');
  button.style.display = 'none';

  const menu = document.createElement('div');
  menu.className = 'fc-autofill-options__menu';
  menu.setAttribute('role', 'menu');
  menu.style.display = 'none';

  const copyItem = document.createElement('button');
  copyItem.type = 'button';
  copyItem.className = 'fc-autofill-options__item';
  copyItem.dataset.fcMode = 'copy';
  copyItem.setAttribute('role', 'menuitemradio');

  const seriesItem = document.createElement('button');
  seriesItem.type = 'button';
  seriesItem.className = 'fc-autofill-options__item';
  seriesItem.dataset.fcMode = 'series';
  seriesItem.setAttribute('role', 'menuitemradio');

  const formattingOnlyItem = document.createElement('button');
  formattingOnlyItem.type = 'button';
  formattingOnlyItem.className = 'fc-autofill-options__item';
  formattingOnlyItem.dataset.fcMode = 'formattingOnly';
  formattingOnlyItem.setAttribute('role', 'menuitemradio');

  const withoutFormattingItem = document.createElement('button');
  withoutFormattingItem.type = 'button';
  withoutFormattingItem.className = 'fc-autofill-options__item';
  withoutFormattingItem.dataset.fcMode = 'withoutFormatting';
  withoutFormattingItem.setAttribute('role', 'menuitemradio');

  const dayItems = (['days', 'weekdays', 'months', 'years'] as const).map((mode) => {
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'fc-autofill-options__item';
    item.dataset.fcMode = mode;
    item.setAttribute('role', 'menuitemradio');
    return item;
  });

  menu.append(copyItem, seriesItem, formattingOnlyItem, withoutFormattingItem, ...dayItems);
  document.body.append(button, menu);

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
    for (let r = detail.src.r0; r <= detail.src.r1; r += 1) {
      for (let c = detail.src.c0; c <= detail.src.c1; c += 1) {
        const key = addrKey({ sheet: detail.src.sheet, row: r, col: c });
        const fmt = state.format.formats.get(key)?.numFmt;
        const cell = state.data.cells.get(key);
        if ((fmt?.kind === 'date' || fmt?.kind === 'datetime') && cell?.value.kind === 'number') {
          return true;
        }
      }
    }
    return false;
  };

  const updateDateItems = (): void => {
    const display = isDateFillCandidate() ? '' : 'none';
    for (const item of dayItems) item.style.display = display;
  };

  const position = (x: number, y: number): void => {
    const left = clamp(x + 6, VIEWPORT_PAD, window.innerWidth - 28);
    const top = clamp(y + 6, VIEWPORT_PAD, window.innerHeight - 28);
    button.style.left = `${left}px`;
    button.style.top = `${top}px`;
    menu.style.left = `${clamp(left, VIEWPORT_PAD, window.innerWidth - 180)}px`;
    menu.style.top = `${clamp(top + 24, VIEWPORT_PAD, window.innerHeight - 80)}px`;
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

  const reapply = (mode: AutoFillMode): void => {
    const current = detail;
    if (!current) return;
    if (history) history.begin();
    let wrote = false;
    try {
      wrote = fillRange(store.getState(), wb, current.src, current.dest, {
        copyOnly: mode === 'copy',
        formatting: formattingModeFor(mode),
        dateUnit: dateUnitFor(mode),
        store,
      });
    } finally {
      if (history) history.end();
    }
    if (wrote) {
      deps.onAfterCommit?.();
      mutators.setActive(store, {
        sheet: current.dest.sheet,
        row: current.dest.r0,
        col: current.dest.c0,
      });
      mutators.extendRangeTo(store, {
        sheet: current.dest.sheet,
        row: current.dest.r1,
        col: current.dest.c1,
      });
      current.mode = mode;
      updateChecked();
    }
    setMenuOpen(false);
    button.focus();
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
  button.addEventListener('click', onButtonClick);
  menu.addEventListener('click', onMenuClick);
  document.addEventListener('mousedown', onDocPointerDown, true);
  document.addEventListener('keydown', onDocKey, true);
  window.addEventListener('scroll', hide, true);

  return {
    detach(): void {
      host.removeEventListener('fc:autofilloptions', onOpen);
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
