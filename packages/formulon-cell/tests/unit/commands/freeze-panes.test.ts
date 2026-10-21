import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { applyFreezeMode, setColsWidth } from '../../../src/commands/row-col-layout.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// A 30 x 18 window scrolled to row 10 (index 9) / column B (index 1).
const scrollTo = (store: SpreadsheetStore, rowStart: number, colStart: number): void => {
  store.setState((s) => ({
    ...s,
    viewport: { ...s.viewport, rowStart, colStart, rowCount: 30, colCount: 18 },
  }));
};

const band = (store: SpreadsheetStore) => {
  const { layout, viewport } = store.getState();
  return {
    rows: layout.freezeRows,
    cols: layout.freezeCols,
    originRow: layout.freezeOriginRow,
    originCol: layout.freezeOriginCol,
    rowStart: viewport.rowStart,
    colStart: viewport.colStart,
  };
};

describe('applyFreezeMode', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
    scrollTo(store, 9, 1);
  });

  it('freezes from the scroll origin up to the active cell', () => {
    mutators.setActive(store, { sheet: 0, row: 14, col: 3 });
    applyFreezeMode(store, 'panes');
    expect(band(store)).toEqual({
      rows: 5,
      cols: 2,
      originRow: 9,
      originCol: 1,
      rowStart: 14,
      colStart: 3,
    });
  });

  it('splits the window in half when the active cell is the top-left', () => {
    mutators.setActive(store, { sheet: 0, row: 9, col: 1 });
    applyFreezeMode(store, 'panes');
    expect(band(store)).toMatchObject({ rows: 15, cols: 9, originRow: 9, originCol: 1 });
  });

  it('splits an unscrolled window at A1', () => {
    scrollTo(store, 0, 0);
    applyFreezeMode(store, 'panes');
    expect(band(store)).toMatchObject({ rows: 15, cols: 9, originRow: 0, originCol: 0 });
  });

  it('freezes rows only in the origin column and columns only in the origin row', () => {
    mutators.setActive(store, { sheet: 0, row: 13, col: 1 });
    applyFreezeMode(store, 'panes');
    expect(band(store)).toMatchObject({ rows: 4, cols: 0 });

    applyFreezeMode(store, 'none');
    scrollTo(store, 9, 1);
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    applyFreezeMode(store, 'panes');
    expect(band(store)).toMatchObject({ rows: 0, cols: 3 });
  });

  it('scrolls an off-screen active cell to the top-left before splitting', () => {
    scrollTo(store, 9, 2);
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    applyFreezeMode(store, 'panes');
    expect(band(store)).toMatchObject({ rows: 15, cols: 9, originRow: 0, originCol: 0 });
  });

  it('freezes the top row and the first column at the visible origin', () => {
    applyFreezeMode(store, 'topRow');
    expect(band(store)).toMatchObject({ rows: 1, cols: 0, originRow: 9, rowStart: 10 });

    applyFreezeMode(store, 'firstColumn');
    expect(band(store)).toMatchObject({ rows: 0, cols: 1, originCol: 1, colStart: 2 });
  });

  it('unfreezes without moving the visible top-left', () => {
    mutators.setActive(store, { sheet: 0, row: 14, col: 3 });
    applyFreezeMode(store, 'panes');
    applyFreezeMode(store, 'none');
    expect(band(store)).toEqual({
      rows: 0,
      cols: 0,
      originRow: 0,
      originCol: 0,
      rowStart: 9,
      colStart: 1,
    });
  });

  it('keeps the body from scrolling back into the band', () => {
    mutators.setActive(store, { sheet: 0, row: 14, col: 3 });
    applyFreezeMode(store, 'panes');
    mutators.scrollBy(store, -100, -100);
    expect(band(store)).toMatchObject({ rowStart: 14, colStart: 3 });
  });

  it('is not an undo step and survives undoing an earlier layout edit', () => {
    const history = new History();
    setColsWidth(store, history, 0, 0, 120);
    mutators.setActive(store, { sheet: 0, row: 14, col: 3 });
    applyFreezeMode(store, 'panes');
    history.undo();
    expect(band(store)).toMatchObject({ rows: 5, cols: 2, originRow: 9, originCol: 1 });
    expect(history.canUndo()).toBe(false);
  });
});
