import { describe, expect, it } from 'vitest';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('setSheetIndex view memo', () => {
  it('starts a sheet visited for the first time at A1, unscrolled', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 5, col: 4 });
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 20, colStart: 3 } }));

    mutators.setSheetIndex(store, 1);

    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 1, row: 0, col: 0 });
    expect([s.viewport.rowStart, s.viewport.colStart]).toEqual([0, 0]);
  });

  it('restores selection, scroll and frozen origin when returning', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 2, 1, { row: 4, col: 2 });
    mutators.setActive(store, { sheet: 0, row: 9, col: 5 });
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 20, colStart: 8 } }));

    mutators.setSheetIndex(store, 1);
    expect(store.getState().layout.freezeOriginRow).toBe(0);
    mutators.setActive(store, { sheet: 1, row: 1, col: 1 });
    mutators.setSheetIndex(store, 0);

    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 9, col: 5 });
    expect([s.viewport.rowStart, s.viewport.colStart]).toEqual([20, 8]);
    expect([s.layout.freezeOriginRow, s.layout.freezeOriginCol]).toEqual([4, 2]);

    mutators.setSheetIndex(store, 1);
    expect(store.getState().selection.active).toEqual({ sheet: 1, row: 1, col: 1 });
  });
});
