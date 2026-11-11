import { describe, expect, it } from 'vitest';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { firstVisibleCol, firstVisibleRow } from '../../../src/store/visible-axis.js';

const scrolled = (rowStart: number, colStart: number) => {
  const store = createSpreadsheetStore();
  store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart, colStart } }));
  return store;
};

describe('store/visible-axis: the first row and column on screen', () => {
  it('is the scrolling pane origin with no frozen panes', () => {
    const store = scrolled(9, 4);
    expect(firstVisibleCol(store.getState())).toBe(4);
    expect(firstVisibleRow(store.getState())).toBe(9);
  });

  it('is the frozen band when panes are frozen', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 2, 2);
    store.setState((s) => ({ ...s, viewport: { ...s.viewport, rowStart: 9, colStart: 5 } }));
    expect(firstVisibleCol(store.getState())).toBe(0);
    expect(firstVisibleRow(store.getState())).toBe(0);
  });

  it('skips hidden columns and rows', () => {
    const store = scrolled(3, 4);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([4, 5]), hiddenRows: new Set([3]) },
    }));
    expect(firstVisibleCol(store.getState())).toBe(6);
    expect(firstVisibleRow(store.getState())).toBe(4);
  });

  it('falls back to the scrolling pane when the whole frozen band is hidden', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 0, 1);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([0]) },
      viewport: { ...s.viewport, colStart: 7 },
    }));
    expect(firstVisibleCol(store.getState())).toBe(7);
  });
});

describe('store/selection-mutators: header selections place the active cell', () => {
  it('puts the active cell of a whole row or column where asked, clamped into the range', () => {
    const store = createSpreadsheetStore();
    mutators.selectRow(store, 2, 4);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 2, col: 4 });
    mutators.selectCols(store, 1, 3, 9);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 9, col: 1 });
    mutators.selectCol(store, 2);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 2 });
  });
});
