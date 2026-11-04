import { describe, expect, it } from 'vitest';
import {
  activeFilterRanges,
  filterColumnActive,
  filterRecordCounts,
} from '../../../src/commands/filter-state.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 1 };

const filtered = () => {
  const store = createSpreadsheetStore();
  store.setState((s) => ({
    ...s,
    ui: {
      ...s.ui,
      filterRanges: [range],
      filterCriteria: [{ range, byCol: 1, hiddenValues: ['Y'] }],
    },
    layout: { ...s.layout, hiddenRows: new Set([7, 8, 9]) },
  }));
  return store.getState();
};

describe('filter state', () => {
  it('counts the surviving data rows against the data rows of the range', () => {
    expect(filterRecordCounts(filtered())).toEqual({ found: 6, total: 9 });
  });

  it('reports nothing while no column carries a criterion', () => {
    const s = createSpreadsheetStore().getState();
    expect(filterRecordCounts(s)).toBeNull();
    expect(activeFilterRanges({ ...s, ui: { ...s.ui, filterRanges: [range] } })).toEqual([]);
  });

  it('marks only the filtered column', () => {
    const s = filtered();
    expect(filterColumnActive(s, range, 1)).toBe(true);
    expect(filterColumnActive(s, range, 0)).toBe(false);
  });
});
