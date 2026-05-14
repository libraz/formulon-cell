import { describe, expect, it } from 'vitest';

import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('store/viewport — setZoom', () => {
  it('clamps zoom to [0.5, 4]', () => {
    const store = createSpreadsheetStore();
    mutators.setZoom(store, 0.1);
    expect(store.getState().viewport.zoom).toBe(0.5);
    mutators.setZoom(store, 10);
    expect(store.getState().viewport.zoom).toBe(4);
  });

  it('passes through valid zoom values', () => {
    const store = createSpreadsheetStore();
    mutators.setZoom(store, 1.25);
    expect(store.getState().viewport.zoom).toBe(1.25);
  });
});

describe('store/viewport — setViewportSize', () => {
  it('updates rowCount + colCount when changed', () => {
    const store = createSpreadsheetStore();
    mutators.setViewportSize(store, 40, 15);
    const v = store.getState().viewport;
    expect(v.rowCount).toBe(40);
    expect(v.colCount).toBe(15);
  });

  it('clamps to a minimum of 1 row / 1 col', () => {
    const store = createSpreadsheetStore();
    mutators.setViewportSize(store, 0, 0);
    const v = store.getState().viewport;
    expect(v.rowCount).toBe(1);
    expect(v.colCount).toBe(1);
  });

  it('floors fractional row/col counts', () => {
    const store = createSpreadsheetStore();
    mutators.setViewportSize(store, 3.9, 2.7);
    const v = store.getState().viewport;
    expect(v.rowCount).toBe(3);
    expect(v.colCount).toBe(2);
  });

  it('no-ops when size is unchanged (reference equality)', () => {
    const store = createSpreadsheetStore();
    mutators.setViewportSize(store, 30, 10);
    const before = store.getState();
    mutators.setViewportSize(store, 30, 10);
    expect(store.getState()).toBe(before);
  });

  it('respects freeze offsets when clamping rowStart/colStart', () => {
    const store = createSpreadsheetStore();
    // Push the viewport row start past the freeze count.
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, freezeRows: 2, freezeCols: 1 },
      viewport: { ...s.viewport, rowStart: 100, colStart: 50 },
    }));
    mutators.setViewportSize(store, 30, 10);
    const v = store.getState().viewport;
    expect(v.rowStart).toBeGreaterThanOrEqual(2);
    expect(v.colStart).toBeGreaterThanOrEqual(1);
  });
});
