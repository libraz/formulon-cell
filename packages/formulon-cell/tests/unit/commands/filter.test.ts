import { beforeEach, describe, expect, it } from 'vitest';
import { applyFilter, clearFilter, setAutoFilter } from '../../../src/commands/filter.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const seedNumber = (store: SpreadsheetStore, row: number, col: number, value: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('filter commands', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('applyFilter hides rows that fail the predicate (header row preserved)', () => {
    // 4-row range: row 0 = header, rows 1..3 = data 10/20/30
    seedNumber(store, 0, 0, 0);
    seedNumber(store, 1, 0, 10);
    seedNumber(store, 2, 0, 20);
    seedNumber(store, 3, 0, 30);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    const hidden = applyFilter(store.getState(), store, range, 0, (cell) => {
      const v = cell?.value as { kind: string; value: number } | undefined;
      return v?.kind === 'number' && v.value >= 20;
    });
    expect(hidden).toBe(1);
    const s = store.getState();
    expect(s.layout.hiddenRows.has(1)).toBe(true);
    expect(s.layout.hiddenRows.has(2)).toBe(false);
    expect(s.layout.hiddenRows.has(3)).toBe(false);
    // Header row never hides.
    expect(s.layout.hiddenRows.has(0)).toBe(false);
  });

  it('applyFilter stamps ui.filterRange so headers can paint chevrons', () => {
    seedNumber(store, 1, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 };
    applyFilter(store.getState(), store, range, 0, () => true);
    expect(store.getState().ui.filterRange).toEqual(range);
  });

  it('setAutoFilter stamps the range without filtering any rows', () => {
    seedNumber(store, 1, 0, 1);
    seedNumber(store, 2, 0, 2);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
    setAutoFilter(store, range);
    const s = store.getState();
    expect(s.ui.filterRange).toEqual(range);
    expect(s.layout.hiddenRows.size).toBe(0);
  });

  it('clearFilter() with no range clears all hidden rows AND filterRange', () => {
    seedNumber(store, 1, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    applyFilter(store.getState(), store, range, 0, () => false);
    expect(store.getState().layout.hiddenRows.size).toBeGreaterThan(0);
    clearFilter(store.getState(), store);
    const s = store.getState();
    expect(s.layout.hiddenRows.size).toBe(0);
    expect(s.ui.filterRange).toBeNull();
  });

  it('clearFilter(range) clears filterRange only when the range matches', () => {
    seedNumber(store, 1, 0, 1);
    const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 };
    applyFilter(store.getState(), store, range, 0, () => true);
    // Same range — clears.
    clearFilter(store.getState(), store, range);
    expect(store.getState().ui.filterRange).toBeNull();

    // Re-apply, then clear with a different range — filterRange survives.
    applyFilter(store.getState(), store, range, 0, () => true);
    clearFilter(store.getState(), store, { sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 });
    expect(store.getState().ui.filterRange).toEqual(range);
  });
});
