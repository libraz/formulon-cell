import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const seedAndMirror = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const fireClipboard = (
  host: HTMLElement,
  type: 'copy' | 'cut' | 'paste',
  data?: string,
): { event: ClipboardEvent; transfer: DataTransfer } => {
  const transfer = new DataTransfer();
  if (data !== undefined) transfer.setData('text/plain', data);
  const event = new ClipboardEvent(type, {
    clipboardData: transfer as unknown as DataTransfer,
    bubbles: true,
    cancelable: true,
  });
  host.dispatchEvent(event);
  return { event, transfer };
};

describe('attachClipboard', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: ReturnType<typeof vi.fn>;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('copy writes TSV to the clipboard and captures a snapshot', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 1 },
      { row: 0, col: 1, value: 'two' },
    ]);
    setRange(store, 0, 0, 0, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event, transfer } = fireClipboard(host, 'copy');
    expect(transfer.getData('text/plain')).toBe('1\ttwo');
    expect(event.defaultPrevented).toBe(true);
    expect(handle.getSnapshot()).not.toBeNull();
    expect(handle.getSnapshot()?.rows).toBe(1);
    expect(handle.getSnapshot()?.cols).toBe(2);
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it('copy is a no-op while the editor is active', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'edit', raw: 'x', caret: 0 } },
    }));
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event, transfer } = fireClipboard(host, 'copy');
    expect(transfer.getData('text/plain')).toBe('');
    expect(event.defaultPrevented).toBe(false);
    expect(handle.getSnapshot()).toBeNull();
    handle.detach();
  });

  it('cut writes TSV, snapshots, blanks the source, and notifies onAfterCommit', () => {
    seedAndMirror(store, wb, [
      { row: 0, col: 0, value: 5 },
      { row: 0, col: 1, value: 6 },
    ]);
    setRange(store, 0, 0, 0, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event, transfer } = fireClipboard(host, 'cut');
    expect(transfer.getData('text/plain')).toBe('5\t6');
    expect(event.defaultPrevented).toBe(true);
    expect(handle.getSnapshot()).not.toBeNull();
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('cut is a no-op while the editor is active', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 5 }]);
    setRange(store, 0, 0, 0, 0);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'enter', raw: 'x' } },
    }));
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'cut');
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    handle.detach();
  });

  it('paste reads TSV and writes through pasteTSV, calling onAfterCommit', () => {
    setRange(store, 1, 1, 1, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event } = fireClipboard(host, 'paste', 'foo\t42');
    expect(event.defaultPrevented).toBe(true);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'foo' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'number', value: 42 });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('paste with empty payload is a no-op', () => {
    setRange(store, 1, 1, 1, 1);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    const { event } = fireClipboard(host, 'paste', '');
    // Empty payload short-circuits before preventDefault.
    expect(event.defaultPrevented).toBe(false);
    expect(onAfterCommit).not.toHaveBeenCalled();
    handle.detach();
  });

  it('paste is a no-op while the editor is active', () => {
    setRange(store, 1, 1, 1, 1);
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, editor: { kind: 'edit', raw: '', caret: 0 } },
    }));
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'paste', 'foo\t42');
    expect(onAfterCommit).not.toHaveBeenCalled();
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    handle.detach();
  });

  it('detach removes listeners so subsequent events do not fire handlers', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });
    handle.detach();

    const { event } = fireClipboard(host, 'copy');
    expect(event.defaultPrevented).toBe(false);
    expect(handle.getSnapshot()).toBeNull();
  });

  it('copy without clipboardData on the event short-circuits gracefully', () => {
    seedAndMirror(store, wb, [{ row: 0, col: 0, value: 1 }]);
    setRange(store, 0, 0, 0, 0);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    // Construct a copy event without clipboardData.
    const event = new ClipboardEvent('copy', { bubbles: true, cancelable: true });
    host.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    handle.detach();
  });
});
