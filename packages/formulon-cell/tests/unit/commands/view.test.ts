import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  clearSheetBackgroundImage,
  setGridlinesVisible,
  setHeadingsVisible,
  setR1C1ReferenceStyle,
  setSheetBackgroundImage,
  setShowFormulas,
  setStatusAggregates,
  setWorkbookView,
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

  it('sets the workbook view mode surfaced by View > Workbook Views', () => {
    const store = createSpreadsheetStore();

    expect(store.getState().ui.workbookView).toBe('normal');
    setWorkbookView(store, 'pageLayout');
    expect(store.getState().ui.workbookView).toBe('pageLayout');
    setWorkbookView(store, 'pageBreakPreview');
    expect(store.getState().ui.workbookView).toBe('pageBreakPreview');
  });

  it('sets and clears sheet background image URLs', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    setSheetBackgroundImage(store, 0, ' https://example.test/bg.png ', history);
    expect(store.getState().ui.sheetBackgroundImages.get(0)).toBe('https://example.test/bg.png');

    expect(history.undo()).toBe(true);
    expect(store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(history.redo()).toBe(true);
    expect(store.getState().ui.sheetBackgroundImages.get(0)).toBe('https://example.test/bg.png');

    clearSheetBackgroundImage(store, 0, history);
    expect(store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(history.undo()).toBe(true);
    expect(store.getState().ui.sheetBackgroundImages.get(0)).toBe('https://example.test/bg.png');
  });

  it('skips sheet background history when the snapshot does not change', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    clearSheetBackgroundImage(store, 0, history);
    expect(history.canUndo()).toBe(false);

    setSheetBackgroundImage(store, 0, 'https://example.test/bg.png', history);
    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);

    setSheetBackgroundImage(store, 0, 'https://example.test/bg.png', history);
    setSheetBackgroundImage(store, 0, ' https://example.test/bg.png ', history);
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().ui.sheetBackgroundImages.has(0)).toBe(false);
    expect(history.canUndo()).toBe(false);
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
