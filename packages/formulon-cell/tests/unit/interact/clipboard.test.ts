import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachClipboard } from '../../../src/interact/clipboard.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

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
  let onAfterCommit: Mock<() => void>;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
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
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
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

  it('copy clears a stale copy marquee when the next copy cannot be materialized', () => {
    mutators.setCopyRanges(store, [
      { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 16383 },
      { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 16383 },
    ]);
    setRange(store, 0, 0, 1_048_575, 16_383);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'copy');

    expect(store.getState().ui.copyRange).toBeNull();
    expect(store.getState().ui.copyRanges).toBeNull();
    handle.detach();
  });

  it('replaces row copy marquees when copying a column next', () => {
    seedAndMirror(store, wb, [
      { row: 2, col: 1, value: 'row' },
      { row: 0, col: 3, value: 'col' },
    ]);
    setRange(store, 2, 0, 2, 16383);
    const handle = attachClipboard({ host, store, wb, onAfterCommit });

    fireClipboard(host, 'copy');
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 16383 });

    setRange(store, 0, 3, 1048575, 3);
    fireClipboard(host, 'copy');

    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 3, r1: 1048575, c1: 3 });
    expect(store.getState().ui.copyRanges).toBeNull();
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
    expect(store.getState().ui.copyRange).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
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
    expect(store.getState().ui.copyRange).toBeNull();
    expect(store.getState().selection.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
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
