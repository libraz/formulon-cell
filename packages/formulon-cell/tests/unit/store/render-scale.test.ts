import { describe, expect, it } from 'vitest';

import { renderScale } from '../../../src/store/render-scale.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const ROW_MAX_PX = (409 * 96) / 72;

const withScale = (screenScale?: number) => {
  const store = createSpreadsheetStore();
  if (screenScale !== undefined) {
    store.setState((s) => ({ ...s, layout: { ...s.layout, screenScale } }));
  }
  return store;
};

describe('store/renderScale', () => {
  it('returns zoom when screenScale is absent', () => {
    const { layout } = createSpreadsheetStore().getState();
    expect(renderScale(layout, { zoom: 1.5 })).toBe(1.5);
  });

  it('multiplies zoom by screenScale', () => {
    const { layout } = withScale(0.75).getState();
    expect(renderScale(layout, { zoom: 2 })).toBe(1.5);
  });

  it('treats a zero or non-finite zoom as 1', () => {
    const { layout } = withScale(0.75).getState();
    expect(renderScale(layout, { zoom: 0 })).toBe(0.75);
    expect(renderScale(layout, { zoom: Number.NaN })).toBe(0.75);
    expect(renderScale(layout, { zoom: -1 })).toBe(0.75);
    expect(renderScale(layout, { zoom: Number.POSITIVE_INFINITY })).toBe(0.75);
  });

  it('returns screenScale without a viewport', () => {
    expect(renderScale(withScale(0.75).getState().layout)).toBe(0.75);
    expect(renderScale(createSpreadsheetStore().getState().layout)).toBe(1);
  });
});

describe('store/setColWidth + setRowHeight — whole screen px', () => {
  it('rounds to whole px at screenScale 1', () => {
    const store = withScale();
    mutators.setColWidth(store, 0, 30.4);
    mutators.setRowHeight(store, 0, 30.6);
    expect(store.getState().layout.colWidths.get(0)).toBe(30);
    expect(store.getState().layout.rowHeights.get(0)).toBe(31);
  });

  it('rounds to whole screen px at screenScale 0.75', () => {
    const store = withScale(0.75);
    // 30.4 screen px = 40.5333 canonical px; snaps to 30 screen px = 40 canonical.
    mutators.setColWidth(store, 0, 30.4 / 0.75);
    mutators.setRowHeight(store, 0, 30.6 / 0.75);
    expect(store.getState().layout.colWidths.get(0)).toBe(40);
    expect(store.getState().layout.rowHeights.get(0)).toBe(124 / 3);
  });

  it('keeps the existing minimums', () => {
    const store = withScale(0.75);
    mutators.setColWidth(store, 0, 1);
    mutators.setRowHeight(store, 0, 1);
    expect(store.getState().layout.colWidths.get(0)).toBe(28);
    expect(store.getState().layout.rowHeights.get(0)).toBe(16);
  });

  it('caps row height at 409 pt', () => {
    const store = withScale();
    mutators.setRowHeight(store, 0, 2000);
    expect(store.getState().layout.rowHeights.get(0)).toBe(ROW_MAX_PX);
    mutators.setRowHeight(store, 1, 500);
    expect(store.getState().layout.rowHeights.get(1)).toBe(500);
  });

  it('caps row height at 409 pt under screenScale 0.75', () => {
    const store = withScale(0.75);
    mutators.setRowHeight(store, 0, 2000);
    expect(store.getState().layout.rowHeights.get(0)).toBe(ROW_MAX_PX);
  });
});
