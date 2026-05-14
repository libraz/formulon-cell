import { beforeEach, describe, expect, it, vi } from 'vitest';

import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createHostShortcutHandler } from '../../../src/mount/host-shortcuts.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type Handler = ReturnType<typeof createHostShortcutHandler>;

function fakeOpenable() {
  return { open: vi.fn() };
}
function fakePainter() {
  return { activate: vi.fn() };
}

interface Setup {
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
    pasteSpecialDialog: ReturnType<typeof fakeOpenable> | null;
    quickAnalysis: ReturnType<typeof fakeOpenable> | null;
  };
}

function makeSetup(): Setup {
  const recalc = vi.fn();
  const cells = vi.fn().mockReturnValue([]);
  const invalidate = vi.fn();
  const setNumber = vi.fn();
  const wb = {
    recalc,
    cells,
    setNumber,
  } as unknown as WorkbookHandle;
  const store = createSpreadsheetStore();
  const history = new History();
  const hostTag = document.createElement('input');

  const feature = {
    findReplace: fakeOpenable(),
    formatDialog: fakeOpenable(),
    formatPainter: fakePainter(),
    goToDialog: fakeOpenable(),
    hyperlinkDialog: fakeOpenable(),
    pasteSpecialDialog: fakeOpenable(),
    quickAnalysis: fakeOpenable(),
  };

  const handler = createHostShortcutHandler({
    findReplace: () => feature.findReplace,
    formatDialog: () => feature.formatDialog,
    formatPainter: () => feature.formatPainter,
    goToDialog: () => feature.goToDialog,
    history,
    hostTag,
    hyperlinkDialog: () => feature.hyperlinkDialog,
    invalidate,
    pasteSpecialDialog: () => feature.pasteSpecialDialog,
    quickAnalysis: () => feature.quickAnalysis,
    store,
    wb: () => wb,
  });

  return {
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

  it('Ctrl+A selects all', () => {
    s.handler(key({ key: 'a', ctrlKey: true }));
    const sel = s.store.getState().selection.range;
    // selectAll spans the whole canonical sheet rect.
    expect(sel.r0).toBe(0);
    expect(sel.c0).toBe(0);
    expect(sel.r1).toBeGreaterThan(0);
    expect(sel.c1).toBeGreaterThan(0);
  });

  it('Ctrl+1 opens the format dialog', () => {
    s.handler(key({ key: '1', ctrlKey: true }));
    expect(s.feature.formatDialog?.open).toHaveBeenCalledTimes(1);
  });

  it('Ctrl+Shift+V opens paste-special', () => {
    s.handler(key({ key: 'v', ctrlKey: true, shiftKey: true }));
    expect(s.feature.pasteSpecialDialog?.open).toHaveBeenCalledTimes(1);
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

  it('does nothing when the matching feature is unavailable', () => {
    s.feature.findReplace = null;
    s.handler(key({ key: 'f', ctrlKey: true }));
    // No exception thrown; no other side effects.
    expect(s.recalc).not.toHaveBeenCalled();
  });
});
