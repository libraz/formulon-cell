import {
  type PasteSpecialOptions,
  type PasteSpecialResult,
  pasteSpecial,
} from '../commands/clipboard/paste-special.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { type History, recordFormatChange } from '../commands/history.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { rangeRects } from '../render/geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import {
  createFloatingOptionsButton,
  createFloatingOptionsMenuItem,
} from './floating-options-menu.js';
import { clampPanelToViewport } from './overlay-position.js';

export type PasteOptionsMode = 'source' | 'values' | 'formatting';

export interface PasteOptionsActivation {
  source: ClipboardSnapshot;
  before: ClipboardSnapshot;
  range: Range;
}

export interface PasteOptionsDeps {
  host: HTMLElement;
  grid: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  strings: Strings;
  history?: History | null;
  onAfterCommit?: () => void;
}

export interface PasteOptionsHandle {
  detach(): void;
  setStrings(next: Strings): void;
  show(next: PasteOptionsActivation): void;
}

const VIEWPORT_PAD = 4;

const defaultOptions = (what: PasteSpecialOptions['what']): PasteSpecialOptions => ({
  what,
  operation: 'none',
  skipBlanks: false,
  transpose: false,
});

export function attachPasteOptions(deps: PasteOptionsDeps): PasteOptionsHandle {
  const { host, grid, store, wb } = deps;
  const history = deps.history ?? null;
  if (history) wb.attachHistory(history);
  let strings = deps.strings;
  let activation: PasteOptionsActivation | null = null;
  let menuOpen = false;

  const button = createFloatingOptionsButton({ className: 'fc-paste-options__button' });

  const menu = document.createElement('div');
  menu.className = 'fc-paste-options__menu';
  menu.setAttribute('role', 'menu');
  menu.style.display = 'none';

  const sourceItem = makeItem('source');
  const valuesItem = makeItem('values');
  const formattingItem = makeItem('formatting');
  menu.append(sourceItem, valuesItem, formattingItem);
  document.body.append(button, menu);

  const applyLabels = (): void => {
    const t = strings.pasteOptions;
    button.title = t.title;
    button.setAttribute('aria-label', t.title);
    menu.setAttribute('aria-label', t.title);
    sourceItem.textContent = t.keepSourceFormatting;
    valuesItem.textContent = t.values;
    formattingItem.textContent = t.formattingOnly;
  };

  const setMenuOpen = (open: boolean): void => {
    menuOpen = open;
    button.setAttribute('aria-expanded', open ? 'true' : 'false');
    menu.style.display = open ? 'block' : 'none';
    if (open) sourceItem.focus({ preventScroll: true });
  };

  const hide = (): void => {
    activation = null;
    button.style.display = 'none';
    setMenuOpen(false);
  };

  const position = (range: Range): void => {
    const state = store.getState();
    const rects = rangeRects(state.layout, state.viewport, range);
    const hostRect = grid.getBoundingClientRect();
    const anchor = rects[rects.length - 1];
    const x = anchor ? hostRect.left + anchor.x + anchor.w : hostRect.left + 24;
    const y = anchor ? hostRect.top + anchor.y + anchor.h : hostRect.top + 24;
    const { x: left, y: top } = clampPanelToViewport(button, x + 3, y + 3, {
      pad: VIEWPORT_PAD,
      fallbackWidth: 28,
      fallbackHeight: 28,
    });
    button.style.left = `${left}px`;
    button.style.top = `${top}px`;
    const menuPos = clampPanelToViewport(menu, left, top + 24, {
      pad: VIEWPORT_PAD,
      fallbackWidth: 220,
      fallbackHeight: 112,
    });
    menu.style.left = `${menuPos.x}px`;
    menu.style.top = `${menuPos.y}px`;
  };

  const runPaste = (snap: ClipboardSnapshot, what: PasteSpecialOptions['what']) =>
    pasteSpecial(store.getState(), store, wb, snap, defaultOptions(what));

  const applyMode = (mode: PasteOptionsMode): void => {
    const current = activation;
    if (!current) return;
    let result: PasteSpecialResult | null = null;
    const apply = (): void => {
      mutators.setRange(store, current.range);
      if (mode === 'source') {
        result = runPaste(current.source, 'all');
      } else if (mode === 'values') {
        result = runPaste(current.source, 'values');
        mutators.setRange(store, current.range);
        runPaste(current.before, 'formats');
      } else {
        result = runPaste(current.before, 'all');
        mutators.setRange(store, current.range);
        runPaste(current.source, 'formats');
      }
    };
    if (history) {
      history.begin();
      try {
        recordFormatChange(history, store, apply);
      } finally {
        history.end();
      }
    } else {
      apply();
    }
    if (result) deps.onAfterCommit?.();
    hide();
  };

  const onHostShow = (e: Event): void => {
    const detail = (e as CustomEvent<PasteOptionsActivation>).detail;
    if (!detail) return;
    show(detail);
  };

  const show = (next: PasteOptionsActivation): void => {
    activation = {
      source: next.source,
      before: next.before,
      range: { ...next.range },
    };
    applyLabels();
    position(next.range);
    button.style.display = 'block';
    setMenuOpen(false);
  };

  const onButtonClick = (e: MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    if (!activation) return;
    setMenuOpen(!menuOpen);
  };

  const onMenuClick = (e: MouseEvent): void => {
    const target = (e.target as HTMLElement | null)?.closest<HTMLButtonElement>(
      '.fc-paste-options__item',
    );
    const mode = target?.dataset.fcMode as PasteOptionsMode | undefined;
    if (!mode) return;
    e.preventDefault();
    e.stopPropagation();
    applyMode(mode);
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
  host.addEventListener('fc:pasteoptions', onHostShow);
  button.addEventListener('click', onButtonClick);
  menu.addEventListener('click', onMenuClick);
  document.addEventListener('mousedown', onDocPointerDown, true);
  document.addEventListener('keydown', onDocKey, true);
  window.addEventListener('scroll', hide, true);

  return {
    show,
    detach(): void {
      host.removeEventListener('fc:pasteoptions', onHostShow);
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

function makeItem(mode: PasteOptionsMode): HTMLButtonElement {
  return createFloatingOptionsMenuItem({ className: 'fc-paste-options__item', mode });
}
