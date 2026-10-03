import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../../src/commands/history.js';
import {
  autofitColsWidth,
  autofitRowsHeight,
  hiddenInSelection,
  hideCols,
  hideRows,
  setColsWidth,
  setRowsHeight,
  showCols,
  showColsAroundSelection,
  showRows,
  showRowsAroundSelection,
} from '../../../../src/commands/row-col-layout.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';

describe('hideRows / showRows / hideCols / showCols', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('hideRows adds a closed range', () => {
    hideRows(store, null, 2, 4);
    expect(Array.from(store.getState().layout.hiddenRows).sort()).toEqual([2, 3, 4]);
  });

  it('hideRows refuses huge row ranges instead of materializing every row', () => {
    hideRows(store, null, 0, 1_048_575);
    expect(store.getState().layout.hiddenRows.size).toBe(0);
  });

  it('showRows clears entries inside the range', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1, 2, 3, 5]) },
    }));
    showRows(store, null, 2, 3);
    expect(Array.from(store.getState().layout.hiddenRows).sort()).toEqual([1, 5]);
  });

  it('showRows clears huge ranges by visiting existing hidden rows only', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1, 250_000, 900_000]) },
    }));

    showRows(store, null, 0, 1_048_575);

    expect(store.getState().layout.hiddenRows.size).toBe(0);
  });

  it('showRowsAroundSelection clears hidden rows directly adjacent to the selection', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set([1, 2, 4, 5, 9]) },
    }));
    showRowsAroundSelection(store, null, 3, 3);
    expect(Array.from(store.getState().layout.hiddenRows).sort()).toEqual([9]);
  });

  it('hideCols / showCols mirror the row variants', () => {
    hideCols(store, null, 0, 1);
    expect(Array.from(store.getState().layout.hiddenCols).sort()).toEqual([0, 1]);
    showCols(store, null, 0, 0);
    expect(Array.from(store.getState().layout.hiddenCols)).toEqual([1]);
  });

  it('showColsAroundSelection clears hidden columns directly adjacent to the selection', () => {
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([0, 2, 3, 5]) },
    }));
    showColsAroundSelection(store, null, 1, 1);
    expect(Array.from(store.getState().layout.hiddenCols).sort()).toEqual([5]);
  });

  it('round-trips through history', () => {
    const h = new History();
    hideRows(store, h, 2, 3);
    expect(store.getState().layout.hiddenRows.size).toBe(2);
    h.undo();
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    h.redo();
    expect(store.getState().layout.hiddenRows.size).toBe(2);
  });

  it('sets row heights and column widths through history', () => {
    const h = new History();
    setRowsHeight(store, h, 1, 2, 32);
    setColsWidth(store, h, 3, 4, 140);

    expect(store.getState().layout.rowHeights.get(1)).toBe(32);
    expect(store.getState().layout.rowHeights.get(2)).toBe(32);
    expect(store.getState().layout.colWidths.get(3)).toBe(140);
    expect(store.getState().layout.colWidths.get(4)).toBe(140);

    h.undo();
    expect(store.getState().layout.colWidths.size).toBe(0);
    h.undo();
    expect(store.getState().layout.rowHeights.size).toBe(0);
  });

  it('refuses huge row height and row autofit ranges', () => {
    setRowsHeight(store, null, 0, 1_048_575, 32);
    autofitRowsHeight(store, null, 0, 1_048_575);

    expect(store.getState().layout.rowHeights.size).toBe(0);
  });

  it('autofits selected rows and columns through history', () => {
    const h = new History();
    store.setState((s) => ({
      ...s,
      data: {
        ...s.data,
        cells: new Map([
          ['0:1:2', { value: { kind: 'text', value: 'wide enough for autofit' }, formula: null }],
          [
            '0:2:2',
            {
              value: { kind: 'text', value: 'first line\nsecond line\nthird line' },
              formula: null,
            },
          ],
        ]),
      },
    }));

    autofitColsWidth(store, h, 2, 2);
    autofitRowsHeight(store, h, 2, 2);

    expect(store.getState().layout.colWidths.get(2)).toBeGreaterThan(100);
    expect(store.getState().layout.rowHeights.get(2)).toBeGreaterThan(40);

    h.undo();
    expect(store.getState().layout.rowHeights.get(2)).toBeUndefined();
    h.undo();
    expect(store.getState().layout.colWidths.get(2)).toBeUndefined();
  });
});

describe('hiddenInSelection', () => {
  it('returns hidden rows inside [a, b]', () => {
    const layout = createSpreadsheetStore().getState().layout;
    const customLayout = { ...layout, hiddenRows: new Set([2, 5, 7]) };
    expect(hiddenInSelection(customLayout, 'row', 1, 6)).toEqual([2, 5]);
    expect(hiddenInSelection(customLayout, 'row', 6, 1)).toEqual([2, 5]); // order-insensitive
  });

  it('returns hidden cols inside [a, b]', () => {
    const layout = createSpreadsheetStore().getState().layout;
    const customLayout = { ...layout, hiddenCols: new Set([0, 3, 9]) };
    expect(hiddenInSelection(customLayout, 'col', 0, 5)).toEqual([0, 3]);
  });

  it('returns empty when nothing is hidden', () => {
    const layout = createSpreadsheetStore().getState().layout;
    expect(hiddenInSelection(layout, 'row', 0, 100)).toEqual([]);
  });

  it('scans hidden metadata rather than the whole selected row span', () => {
    const layout = createSpreadsheetStore().getState().layout;
    const customLayout = { ...layout, hiddenRows: new Set([5, 500_000, 1_000_000]) };

    expect(hiddenInSelection(customLayout, 'row', 0, 1_048_575)).toEqual([5, 500_000, 1_000_000]);
  });
});

describe('hideRows / hideCols engine sync', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('forwards hideRows to setRowHidden when capability is on', () => {
    const calls: { row: number; hidden: boolean }[] = [];
    const wb = {
      capabilities: { hiddenRowsCols: true },
      setRowHidden: (_sheet: number, row: number, hidden: boolean) => {
        calls.push({ row, hidden });
        return true;
      },
    } as unknown as WorkbookHandle;
    const h = new History();
    hideRows(store, h, 2, 4, wb);
    expect(calls).toEqual([
      { row: 2, hidden: true },
      { row: 3, hidden: true },
      { row: 4, hidden: true },
    ]);
    h.undo();
    // After undo, the same rows must be unhidden.
    expect(calls.slice(-3)).toEqual([
      { row: 2, hidden: false },
      { row: 3, hidden: false },
      { row: 4, hidden: false },
    ]);
  });

  it('skips engine calls when hiddenRowsCols capability is off', () => {
    const calls: { row: number; hidden: boolean }[] = [];
    const wb = {
      capabilities: { hiddenRowsCols: false },
      setRowHidden: (_s: number, row: number, hidden: boolean) => {
        calls.push({ row, hidden });
        return false;
      },
    } as unknown as WorkbookHandle;
    hideRows(store, null, 0, 0, wb);
    expect(calls).toEqual([]);
    expect(store.getState().layout.hiddenRows.has(0)).toBe(true);
  });
});
