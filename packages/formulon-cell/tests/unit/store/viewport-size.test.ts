import { describe, expect, it } from 'vitest';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('mutators.setViewportSize', () => {
  it('updates visible row and column counts from the measured grid size', () => {
    const store = createSpreadsheetStore();

    mutators.setViewportSize(store, 34, 22);

    expect(store.getState().viewport.rowCount).toBe(34);
    expect(store.getState().viewport.colCount).toBe(22);
  });

  it('clamps scroll starts when a larger measured viewport reaches sheet bounds', () => {
    const store = createSpreadsheetStore();
    mutators.scrollBy(store, 1_048_575, 16_383);

    mutators.setViewportSize(store, 100, 40);

    expect(store.getState().viewport.rowStart).toBe(1_048_576 - 100);
    expect(store.getState().viewport.colStart).toBe(16_384 - 40);
  });

  it('keeps the body viewport past frozen panes', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 3, 2);
    mutators.setViewportSize(store, 20, 12);

    expect(store.getState().viewport.rowStart).toBe(3);
    expect(store.getState().viewport.colStart).toBe(2);
  });
});
