import { describe, expect, it } from 'vitest';
import {
  setGridlinesVisible,
  setHeadingsVisible,
  setR1C1ReferenceStyle,
  setShowFormulas,
  setStatusAggregates,
  setZoomPercent,
  setZoomScale,
  toggleStatusAggregate,
} from '../../../src/commands/view.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('view commands', () => {
  it('toggles spreadsheet-style view flags', () => {
    const store = createSpreadsheetStore();

    setGridlinesVisible(store, false);
    setHeadingsVisible(store, false);
    setShowFormulas(store, true);
    setR1C1ReferenceStyle(store, true);

    expect(store.getState().ui.showGridLines).toBe(false);
    expect(store.getState().ui.showHeaders).toBe(false);
    expect(store.getState().ui.showFormulas).toBe(true);
    expect(store.getState().ui.r1c1).toBe(true);
  });

  it('sets zoom by scale or spreadsheet-style percent and uses store clamping', () => {
    const store = createSpreadsheetStore();

    setZoomScale(store, 1.25);
    expect(store.getState().viewport.zoom).toBe(1.25);

    setZoomPercent(store, 175);
    expect(store.getState().viewport.zoom).toBe(1.75);

    setZoomPercent(store, 25);
    expect(store.getState().viewport.zoom).toBe(0.5);

    setZoomScale(store, 9);
    expect(store.getState().viewport.zoom).toBe(4);
  });

  it('sets and toggles status bar aggregates', () => {
    const store = createSpreadsheetStore();

    setStatusAggregates(store, ['sum', 'max']);
    expect(store.getState().ui.statusAggs).toEqual(['sum', 'max']);

    toggleStatusAggregate(store, 'sum');
    expect(store.getState().ui.statusAggs).toEqual(['max']);

    toggleStatusAggregate(store, 'average');
    expect(store.getState().ui.statusAggs).toEqual(['max', 'average']);
  });
});
