import { describe, expect, it } from 'vitest';
import { setAutoFilter } from '../../../src/commands/filter.js';
import { activeFilterRange, filterRangeAt } from '../../../src/commands/filter-ranges.js';
import { defaultTableOverlay } from '../../../src/commands/format-as-table.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const tableRange = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 };
const sheetRange = { sheet: 0, r0: 0, c0: 4, r1: 4, c1: 5 };

function storeWithFilters() {
  const store = createSpreadsheetStore();
  mutators.upsertTableOverlay(store, defaultTableOverlay('t', tableRange));
  mutators.setFilterRanges(store, [tableRange, sheetRange]);
  return store;
}

describe('filter range lookups', () => {
  it('finds the filter range holding an address', () => {
    const state = storeWithFilters().getState();
    expect(filterRangeAt(state, { sheet: 0, row: 2, col: 1 })).toEqual(tableRange);
    expect(filterRangeAt(state, { sheet: 0, row: 2, col: 5 })).toEqual(sheetRange);
    expect(filterRangeAt(state, { sheet: 0, row: 2, col: 3 })).toBeNull();
    expect(filterRangeAt(state, { sheet: 1, row: 2, col: 1 })).toBeNull();
  });

  it('picks the table holding the active cell, else the worksheet filter', () => {
    const store = storeWithFilters();
    mutators.setActive(store, { sheet: 0, row: 3, col: 0 });
    expect(activeFilterRange(store.getState())).toEqual(tableRange);
    mutators.setActive(store, { sheet: 0, row: 3, col: 5 });
    expect(activeFilterRange(store.getState())).toEqual(sheetRange);
    mutators.setActive(store, { sheet: 0, row: 8, col: 8 });
    expect(activeFilterRange(store.getState())).toEqual(sheetRange);

    mutators.setFilterRanges(store, [tableRange]);
    expect(activeFilterRange(store.getState())).toBeNull();
  });

  it('replaces only the worksheet filter when another worksheet range is set', () => {
    const store = storeWithFilters();
    const moved = { sheet: 0, r0: 6, c0: 4, r1: 9, c1: 4 };
    setAutoFilter(store, moved);
    expect(store.getState().ui.filterRanges).toEqual([tableRange, moved]);
    // A grown table filter replaces its old extent.
    const grown = { ...tableRange, r1: 5 };
    mutators.upsertTableOverlay(store, defaultTableOverlay('t', grown));
    setAutoFilter(store, grown);
    expect(store.getState().ui.filterRanges).toEqual([moved, grown]);
  });
});
