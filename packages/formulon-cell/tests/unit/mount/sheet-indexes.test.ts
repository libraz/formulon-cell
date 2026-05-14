import { describe, expect, it } from 'vitest';

import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { hiddenSheetIndexes, visibleSheetIndexes } from '../../../src/mount/sheet-indexes.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

function fakeWb(sheetCount: number): WorkbookHandle {
  return { sheetCount } as unknown as WorkbookHandle;
}

function storeWithHidden(hidden: number[]): SpreadsheetStore {
  const store = createSpreadsheetStore();
  const hiddenSet = new Set(hidden);
  // Mutate the layout slice directly. The store exposes a setState API on its
  // public surface; the field is shared with the rest of the layout slice.
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenSheets: hiddenSet },
  }));
  return store;
}

describe('mount/sheet-indexes', () => {
  it('returns all sheets when none are hidden', () => {
    const store = storeWithHidden([]);
    expect(visibleSheetIndexes(fakeWb(3), store)).toEqual([0, 1, 2]);
    expect(hiddenSheetIndexes(fakeWb(3), store)).toEqual([]);
  });

  it('partitions sheets by hidden membership', () => {
    const store = storeWithHidden([1]);
    expect(visibleSheetIndexes(fakeWb(3), store)).toEqual([0, 2]);
    expect(hiddenSheetIndexes(fakeWb(3), store)).toEqual([1]);
  });

  it('returns visible/hidden in ascending sheet order even if Set is unordered', () => {
    const store = storeWithHidden([3, 0]);
    expect(visibleSheetIndexes(fakeWb(5), store)).toEqual([1, 2, 4]);
    expect(hiddenSheetIndexes(fakeWb(5), store)).toEqual([0, 3]);
  });

  it('handles workbooks with zero sheets', () => {
    const store = storeWithHidden([]);
    expect(visibleSheetIndexes(fakeWb(0), store)).toEqual([]);
    expect(hiddenSheetIndexes(fakeWb(0), store)).toEqual([]);
  });
});
