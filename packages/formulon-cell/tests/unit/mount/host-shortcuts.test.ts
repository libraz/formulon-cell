import { beforeEach, describe, expect, it, vi } from 'vitest';

import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings/en.js';
import { createHostShortcutHandler } from '../../../src/mount/host-shortcuts.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type Handler = ReturnType<typeof createHostShortcutHandler>;

function fakeOpenable() {
  return { open: vi.fn() };
}
function fakePainter() {
  return { activate: vi.fn() };
}

interface Setup {
  addSheet: ReturnType<typeof vi.fn>;
  handler: Handler;
  history: History;
  store: ReturnType<typeof createSpreadsheetStore>;
  wb: WorkbookHandle;
  recalc: ReturnType<typeof vi.fn>;
  cells: ReturnType<typeof vi.fn>;
  invalidate: ReturnType<typeof vi.fn>;
  setNumber: ReturnType<typeof vi.fn>;
  hostTag: HTMLInputElement;
  feature: {
    findReplace: ReturnType<typeof fakeOpenable> | null;
    formatDialog: ReturnType<typeof fakeOpenable> | null;
    formatPainter: ReturnType<typeof fakePainter> | null;
    goToDialog: ReturnType<typeof fakeOpenable> | null;
    hyperlinkDialog: ReturnType<typeof fakeOpenable> | null;
    namedRangeDialog: ReturnType<typeof fakeOpenable> | null;
    pasteSpecialDialog: ReturnType<typeof fakeOpenable> | null;
    quickAnalysis: ReturnType<typeof fakeOpenable> | null;
  };
}

function makeSetup(): Setup {
  const addSheet = vi.fn();
  const recalc = vi.fn();
  const cells = vi.fn().mockReturnValue([]);
  const invalidate = vi.fn();
  const setNumber = vi.fn();
  const wb = {
    capabilities: {},
    recalc,
    cells,
    setNumber,
  } as unknown as WorkbookHandle;
  const store = createSpreadsheetStore();
  const history = new History();
  const host = document.createElement('div');
  const hostTag = document.createElement('input');

  const feature = {
    findReplace: fakeOpenable(),
    formatDialog: fakeOpenable(),
    formatPainter: fakePainter(),
    goToDialog: fakeOpenable(),
    hyperlinkDialog: fakeOpenable(),
    namedRangeDialog: fakeOpenable(),
    pasteSpecialDialog: fakeOpenable(),
    quickAnalysis: fakeOpenable(),
  };

  const handler = createHostShortcutHandler({
    addSheet,
    findReplace: () => feature.findReplace,
    formatDialog: () => feature.formatDialog,
    formatPainter: () => feature.formatPainter,
    goToDialog: () => feature.goToDialog,
    history,
    host,
    hostTag,
    hyperlinkDialog: () => feature.hyperlinkDialog,
    invalidate,
    locale: 'en-US',
    namedRangeDialog: () => feature.namedRangeDialog,
    pasteSpecialDialog: () => feature.pasteSpecialDialog,
    quickAnalysis: () => feature.quickAnalysis,
    store,
    strings: () => en,
    wb: () => wb,
  });

  return {
    addSheet,
    handler,
    history,
    store,
    wb,
    recalc,
    cells,
    invalidate,
    setNumber,
    hostTag,
    feature,
  };
}

function key(over: Partial<KeyboardEventInit & { key: string }>): KeyboardEvent {
  return new KeyboardEvent('keydown', { ...over, cancelable: true });
}

describe('mount/host-shortcuts', () => {
  let s: Setup;
  beforeEach(() => {
    s = makeSetup();
  });

  it('F9 recalcs and invalidates, without requiring meta', () => {
    const e = key({ key: 'F9' });
    s.handler(e);
    expect(s.recalc).toHaveBeenCalledTimes(1);
    expect(s.cells).toHaveBeenCalled();
    expect(s.invalidate).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('Shift+F11 delegates sheet insertion to the sheet-tabs controller', () => {
    const e = key({ key: 'F11', shiftKey: true });
    s.handler(e);
    expect(s.addSheet).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('ignores non-F9 keys when no modifier is held', () => {
    s.handler(key({ key: 'f' }));
    expect(s.feature.findReplace?.open).not.toHaveBeenCalled();
  });

  it('Ctrl+F opens Find & Replace', () => {
    const e = key({ key: 'f', ctrlKey: true });
    s.handler(e);
    expect(s.feature.findReplace?.open).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('Cmd+F also opens Find & Replace (macOS)', () => {
    const e = key({ key: 'f', metaKey: true });
    s.handler(e);
    expect(s.feature.findReplace?.open).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('Ctrl+K opens the hyperlink dialog', () => {
    s.handler(key({ key: 'k', ctrlKey: true }));
    expect(s.feature.hyperlinkDialog?.open).toHaveBeenCalledTimes(1);
  });

  it('leaves Ctrl+A to the grid keyboard handler for two-stage selection', () => {
    const e = key({ key: 'a', ctrlKey: true });
    s.handler(e);
    const sel = s.store.getState().selection.range;
    expect(sel.r0).toBe(0);
    expect(sel.c0).toBe(0);
    expect(sel.r1).toBe(0);
    expect(sel.c1).toBe(0);
    expect(e.defaultPrevented).toBe(false);
  });

  it('Ctrl+1 opens the format dialog', () => {
    s.handler(key({ key: '1', ctrlKey: true }));
    expect(s.feature.formatDialog?.open).toHaveBeenCalledTimes(1);
  });

  it('Ctrl+9/Ctrl+0 hide selected rows and columns, with Shift restoring them', () => {
    s.handler(key({ key: '9', ctrlKey: true }));
    s.handler(key({ key: '0', ctrlKey: true }));
    expect(s.store.getState().layout.hiddenRows.has(0)).toBe(true);
    expect(s.store.getState().layout.hiddenCols.has(0)).toBe(true);

    s.handler(key({ key: '9', ctrlKey: true, shiftKey: true }));
    s.handler(key({ key: '0', ctrlKey: true, shiftKey: true }));
    expect(s.store.getState().layout.hiddenRows.has(0)).toBe(false);
    expect(s.store.getState().layout.hiddenCols.has(0)).toBe(false);
    expect(s.invalidate).toHaveBeenCalledTimes(4);
  });

  it('Ctrl+Shift+L toggles the selected range auto-filter', () => {
    const e = key({ key: 'l', ctrlKey: true, shiftKey: true });
    s.handler(e);
    expect(s.store.getState().ui.filterRange).toEqual(s.store.getState().selection.range);
    expect(e.defaultPrevented).toBe(true);

    s.handler(key({ key: 'l', ctrlKey: true, shiftKey: true }));
    expect(s.store.getState().ui.filterRange).toBeNull();
    expect(s.invalidate).toHaveBeenCalledTimes(2);
  });

  it.each(['t', 'l'])('Ctrl+%s formats the selected range as a table', (keyName) => {
    const e = key({ key: keyName, ctrlKey: true });
    s.handler(e);
    expect(s.store.getState().tables.tables).toHaveLength(1);
    expect(s.store.getState().tables.tables[0]?.range).toEqual(s.store.getState().selection.range);
    expect(e.defaultPrevented).toBe(true);
  });

  it('Ctrl+E routes to the flash-fill command', () => {
    const e = key({ key: 'e', ctrlKey: true });
    s.handler(e);
    expect(e.defaultPrevented).toBe(true);
    expect(s.invalidate).toHaveBeenCalledTimes(1);
  });

  it.each([
    { event: { key: '+', ctrlKey: true, shiftKey: true }, kind: 'insert' },
    { event: { key: '-', ctrlKey: true }, kind: 'delete' },
  ])('opens the $kind Cells dialog from its shortcut', ({ event, kind }) => {
    const e = key(event);
    s.handler(e);
    expect(document.querySelector('.fc-cellshift')).not.toBeNull();
    expect(document.querySelector('.fc-cellshift')?.getAttribute('aria-label')).toBe(
      kind === 'insert' ? en.ribbonMenu.insertCells : en.ribbonMenu.deleteCells,
    );
    expect(e.defaultPrevented).toBe(true);
  });

  it.each([
    { event: { key: '~', code: 'Backquote', ctrlKey: true, shiftKey: true }, expected: 'general' },
    { event: { key: '!', code: 'Digit1', ctrlKey: true, shiftKey: true }, expected: 'fixed' },
    { event: { key: '@', code: 'Digit2', ctrlKey: true, shiftKey: true }, expected: 'time' },
    { event: { key: '#', code: 'Digit3', ctrlKey: true, shiftKey: true }, expected: 'date' },
    { event: { key: '$', code: 'Digit4', ctrlKey: true, shiftKey: true }, expected: 'currency' },
    { event: { key: '%', code: 'Digit5', ctrlKey: true, shiftKey: true }, expected: 'percent' },
    { event: { key: '^', code: 'Digit6', ctrlKey: true, shiftKey: true }, expected: 'scientific' },
  ])('Ctrl+Shift+$key applies the $expected number format', ({ event, expected }) => {
    const e = key(event);
    s.handler(e);
    expect(s.store.getState().ui.pendingFormat?.format.numFmt?.kind).toBe(expected);
    expect(e.defaultPrevented).toBe(true);
  });

  it('F4 repeats a direct number format on the current selection, not the original cell', () => {
    s.handler(key({ key: '$', code: 'Digit4', ctrlKey: true, shiftKey: true }));
    mutators.setActive(s.store, { sheet: 0, row: 0, col: 1 });

    const e = key({ key: 'F4' });
    s.handler(e);

    expect(s.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 1 },
      format: { numFmt: { kind: 'currency', symbol: '$', decimals: 2 } },
    });
    expect(e.defaultPrevented).toBe(true);
  });

  it('F4 reapplies a format-toggle result instead of toggling the new cell', () => {
    s.handler(key({ key: 'b', ctrlKey: true }));
    mutators.setActive(s.store, { sheet: 0, row: 0, col: 1 });

    s.handler(key({ key: 'F4' }));

    expect(s.store.getState().ui.pendingFormat).toEqual({
      addr: { sheet: 0, row: 0, col: 1 },
      format: { bold: true },
    });
  });

  it('Ctrl+Shift+V opens paste-special', () => {
    s.handler(key({ key: 'v', ctrlKey: true, shiftKey: true }));
    expect(s.feature.pasteSpecialDialog?.open).toHaveBeenCalledTimes(1);
  });

  it('Ctrl+Alt+V opens paste-special', () => {
    const e = key({ key: 'v', ctrlKey: true, altKey: true });
    s.handler(e);
    expect(s.feature.pasteSpecialDialog?.open).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('Ctrl+F3 opens Name Manager', () => {
    const e = key({ key: 'F3', ctrlKey: true });
    s.handler(e);
    expect(s.feature.namedRangeDialog?.open).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('Ctrl+Q opens quick analysis (but Cmd+Q does not — reserved for the OS)', () => {
    s.handler(key({ key: 'q', ctrlKey: true }));
    expect(s.feature.quickAnalysis?.open).toHaveBeenCalledTimes(1);
    s.feature.quickAnalysis?.open.mockClear();

    s.handler(key({ key: 'q', metaKey: true }));
    expect(s.feature.quickAnalysis?.open).not.toHaveBeenCalled();
  });

  it('Ctrl+Shift+C activates the format painter (non-sticky)', () => {
    s.handler(key({ key: 'c', ctrlKey: true, shiftKey: true }));
    expect(s.feature.formatPainter?.activate).toHaveBeenCalledWith(false);
  });

  it('Ctrl+` toggles formula view in the store', () => {
    const before = s.store.getState().ui.showFormulas;
    s.handler(key({ key: '`', ctrlKey: true }));
    expect(s.store.getState().ui.showFormulas).toBe(!before);
  });

  it('Ctrl+Alt+R toggles R1C1 mode', () => {
    const before = s.store.getState().ui.r1c1;
    s.handler(key({ key: 'r', ctrlKey: true, altKey: true }));
    expect(s.store.getState().ui.r1c1).toBe(!before);
  });

  it.each([
    { key: ':', shiftKey: false },
    { key: ':', shiftKey: true },
    { key: ';', shiftKey: true },
  ])('Ctrl+%s enters the current time on JIS and US layouts', (event) => {
    const e = key({ ...event, ctrlKey: true });
    s.handler(e);
    expect(s.setNumber).toHaveBeenCalledTimes(1);
    expect(e.defaultPrevented).toBe(true);
  });

  it('does nothing when the matching feature is unavailable', () => {
    s.feature.findReplace = null;
    s.handler(key({ key: 'f', ctrlKey: true }));
    // No exception thrown; no other side effects.
    expect(s.recalc).not.toHaveBeenCalled();
  });
});
