import { describe, expect, it } from 'vitest';

import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('store/selection — mutators', () => {
  it('setActive moves the active cell and collapses the range to it', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 4, col: 3 });
    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 4, col: 3 });
    expect(s.selection.range).toEqual({ sheet: 0, r0: 4, c0: 3, r1: 4, c1: 3 });
  });

  it('extendRangeTo grows the range from anchor toward the target', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.extendRangeTo(store, { sheet: 0, row: 4, col: 5 });
    const r = store.getState().selection.range;
    expect(r).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 4, c1: 5 });
  });

  it('extendRangeTo handles "shrink back" (target above-left of anchor)', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 5, col: 5 });
    mutators.extendRangeTo(store, { sheet: 0, row: 2, col: 2 });
    const r = store.getState().selection.range;
    expect(r.r0).toBe(2);
    expect(r.c0).toBe(2);
    expect(r.r1).toBe(5);
    expect(r.c1).toBe(5);
  });

  it('setRange overrides the selected range without touching active', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
    const s = store.getState();
    expect(s.selection.range).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 3, c1: 3 });
    expect(s.selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
  });

  it('selectRow selects the entire row across the sheet width', () => {
    const store = createSpreadsheetStore();
    mutators.selectRow(store, 7);
    const r = store.getState().selection.range;
    expect(r.r0).toBe(7);
    expect(r.r1).toBe(7);
    // Whole-row selection spans many columns.
    expect(r.c1 - r.c0).toBeGreaterThan(50);
  });

  it('selectCol selects the entire column across the sheet height', () => {
    const store = createSpreadsheetStore();
    mutators.selectCol(store, 2);
    const r = store.getState().selection.range;
    expect(r.c0).toBe(2);
    expect(r.c1).toBe(2);
    expect(r.r1 - r.r0).toBeGreaterThan(50);
  });

  it('selectAll selects every cell on the sheet', () => {
    const store = createSpreadsheetStore();
    mutators.selectAll(store);
    const r = store.getState().selection.range;
    expect(r.r0).toBe(0);
    expect(r.c0).toBe(0);
    expect(r.r1).toBeGreaterThan(50);
    expect(r.c1).toBeGreaterThan(50);
  });

  it('addExtraCell demotes the prior primary range into extraRanges and promotes the new cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.addExtraCell(store, { sheet: 0, row: 5, col: 5 });
    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 5, col: 5 });
    expect(s.selection.extraRanges?.length).toBe(1);
    expect(s.selection.extraRanges?.[0]).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
  });

  it('addExtraCell is a no-op when called on the current active cell', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.addExtraCell(store, { sheet: 0, row: 0, col: 0 });
    const s = store.getState();
    expect(s.selection.extraRanges?.length ?? 0).toBe(0);
  });
});
