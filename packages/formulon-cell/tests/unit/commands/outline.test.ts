import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  colGroupRangeAt,
  collapseColGroup,
  collapseRowGroup,
  expandColGroup,
  expandRowGroup,
  groupCols,
  groupRows,
  isColGroupCollapsed,
  isRowGroupCollapsed,
  MAX_OUTLINE_LEVEL,
  OUTLINE_GUTTER_PER_LEVEL,
  rowGroupRangeAt,
  ungroupCols,
  ungroupRows,
} from '../../../src/commands/outline.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('groupRows / ungroupRows', () => {
  it('groupRows raises level by 1, capped at 7', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 2, 4);
    let s = store.getState();
    expect(s.layout.outlineRows.get(2)).toBe(1);
    expect(s.layout.outlineRows.get(3)).toBe(1);
    expect(s.layout.outlineRows.get(4)).toBe(1);
    expect(s.layout.outlineRows.get(5)).toBeUndefined();

    // Cap at MAX_OUTLINE_LEVEL.
    for (let i = 0; i < 10; i += 1) groupRows(store, null, 3, 3);
    s = store.getState();
    expect(s.layout.outlineRows.get(3)).toBe(MAX_OUTLINE_LEVEL);
  });

  it('groupRows updates outlineRowGutter to max level × per-level slot', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 1, 5);
    expect(store.getState().layout.outlineRowGutter).toBe(OUTLINE_GUTTER_PER_LEVEL);
    groupRows(store, null, 2, 4);
    expect(store.getState().layout.outlineRowGutter).toBe(OUTLINE_GUTTER_PER_LEVEL * 2);
  });

  it('ungroupRows decrements and removes when reaching 0', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 0, 2);
    groupRows(store, null, 1, 1);
    expect(store.getState().layout.outlineRows.get(1)).toBe(2);
    ungroupRows(store, null, 1, 1);
    expect(store.getState().layout.outlineRows.get(1)).toBe(1);
    ungroupRows(store, null, 0, 2);
    expect(store.getState().layout.outlineRows.size).toBe(0);
    expect(store.getState().layout.outlineRowGutter).toBe(0);
  });

  it('ignores empty range (r0 > r1)', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 5, 3);
    expect(store.getState().layout.outlineRows.size).toBe(0);
  });

  it('records a single layout history entry for the group operation', () => {
    const store = createSpreadsheetStore();
    const h = new History();
    groupRows(store, h, 0, 4);
    expect(store.getState().layout.outlineRows.get(2)).toBe(1);
    expect(h.canUndo()).toBe(true);
    h.undo();
    expect(store.getState().layout.outlineRows.size).toBe(0);
    h.redo();
    expect(store.getState().layout.outlineRows.get(2)).toBe(1);
  });
});

describe('groupCols / ungroupCols', () => {
  it('mirrors row behavior for columns', () => {
    const store = createSpreadsheetStore();
    groupCols(store, null, 1, 3);
    expect(store.getState().layout.outlineCols.get(2)).toBe(1);
    expect(store.getState().layout.outlineColGutter).toBe(OUTLINE_GUTTER_PER_LEVEL);
    ungroupCols(store, null, 1, 3);
    expect(store.getState().layout.outlineCols.size).toBe(0);
    expect(store.getState().layout.outlineColGutter).toBe(0);
  });
});

describe('rowGroupRangeAt / colGroupRangeAt', () => {
  it('returns the contiguous run at the given level', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 2, 6);
    const range = rowGroupRangeAt(store.getState().layout, 4, 1);
    expect(range).toEqual({ r0: 2, r1: 6 });
  });

  it('returns null when row is below the requested level', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 2, 4);
    expect(rowGroupRangeAt(store.getState().layout, 0, 1)).toBeNull();
    expect(rowGroupRangeAt(store.getState().layout, 4, 2)).toBeNull();
  });

  it('walks separately for each nested level', () => {
    const store = createSpreadsheetStore();
    // Outer group rows 1..6 at level 1.
    groupRows(store, null, 1, 6);
    // Nested group rows 3..4 at level 2.
    groupRows(store, null, 3, 4);
    const layout = store.getState().layout;
    expect(rowGroupRangeAt(layout, 4, 1)).toEqual({ r0: 1, r1: 6 });
    expect(rowGroupRangeAt(layout, 4, 2)).toEqual({ r0: 3, r1: 4 });
  });

  it('colGroupRangeAt mirrors row behavior', () => {
    const store = createSpreadsheetStore();
    groupCols(store, null, 2, 5);
    expect(colGroupRangeAt(store.getState().layout, 3, 1)).toEqual({ c0: 2, c1: 5 });
  });
});

describe('collapse / expand', () => {
  it('collapseRowGroup adds rows to hiddenRows; expandRowGroup removes them', () => {
    const store = createSpreadsheetStore();
    groupRows(store, null, 1, 3);
    collapseRowGroup(store, null, 1, 3);
    expect(isRowGroupCollapsed(store.getState().layout, 1, 3)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(true);
    expandRowGroup(store, null, 1, 3);
    expect(isRowGroupCollapsed(store.getState().layout, 1, 3)).toBe(false);
    expect(store.getState().layout.hiddenRows.size).toBe(0);
  });

  it('collapseColGroup mirrors row behavior', () => {
    const store = createSpreadsheetStore();
    groupCols(store, null, 2, 4);
    collapseColGroup(store, null, 2, 4);
    expect(isColGroupCollapsed(store.getState().layout, 2, 4)).toBe(true);
    expandColGroup(store, null, 2, 4);
    expect(isColGroupCollapsed(store.getState().layout, 2, 4)).toBe(false);
  });

  it('isRowGroupCollapsed reports true when ANY row in the band is hidden', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => {
      const next = new Set(s.layout.hiddenRows);
      next.add(3);
      return { ...s, layout: { ...s.layout, hiddenRows: next } };
    });
    expect(isRowGroupCollapsed(store.getState().layout, 1, 5)).toBe(true);
    expect(isRowGroupCollapsed(store.getState().layout, 6, 8)).toBe(false);
  });
});

describe('outline engine sync', () => {
  it('groupRows pushes setRowOutline for each row in the band', () => {
    const store = createSpreadsheetStore();
    const calls: { row: number; level: number }[] = [];
    const wb = {
      capabilities: { outlines: true },
      setRowOutline: (_sheet: number, row: number, level: number) => {
        calls.push({ row, level });
        return true;
      },
    } as unknown as WorkbookHandle;
    groupRows(store, null, 0, 2, wb);
    expect(calls).toEqual([
      { row: 0, level: 1 },
      { row: 1, level: 1 },
      { row: 2, level: 1 },
    ]);
  });

  it('ungroupCols emits level=0 when the entry drops to zero', () => {
    const store = createSpreadsheetStore();
    const noopWb = {
      capabilities: { outlines: true },
      setColumnOutline: () => true,
    } as unknown as WorkbookHandle;
    groupCols(store, null, 1, 1, noopWb); // get outlineCols[1] = 1

    const calls: { first: number; last: number; level: number }[] = [];
    const wb = {
      capabilities: { outlines: true },
      setColumnOutline: (_sheet: number, first: number, last: number, level: number) => {
        calls.push({ first, last, level });
        return true;
      },
    } as unknown as WorkbookHandle;
    ungroupCols(store, null, 1, 1, wb);
    expect(calls).toEqual([{ first: 1, last: 1, level: 0 }]);
  });

  it('collapseRowGroup forwards hidden=true to the engine', () => {
    const store = createSpreadsheetStore();
    const calls: { row: number; hidden: boolean }[] = [];
    const wb = {
      capabilities: { hiddenRowsCols: true, outlines: true },
      setRowHidden: (_sheet: number, row: number, hidden: boolean) => {
        calls.push({ row, hidden });
        return true;
      },
      setRowOutline: () => true,
    } as unknown as WorkbookHandle;
    collapseRowGroup(store, null, 3, 4, wb);
    expect(calls).toEqual([
      { row: 3, hidden: true },
      { row: 4, hidden: true },
    ]);
  });
});
