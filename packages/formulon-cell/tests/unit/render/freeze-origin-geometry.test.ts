import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  cellRect,
  frozenColsWidth,
  frozenRowsHeight,
  hitTest,
  isColVisible,
  isRowVisible,
  rangeRects,
  type ViewLayout,
  visibleCols,
  visibleRows,
} from '../../../src/render/geometry.js';
import type { ViewportSlice } from '../../../src/store/store.js';

// Rows 9..11 and columns 1..2 are frozen; the body starts at row 12 / column 3.
const layout: ViewLayout = {
  rtl: false,
  viewWidth: 0,
  colWidths: new Map(),
  rowHeights: new Map(),
  defaultColWidth: 100,
  defaultRowHeight: 20,
  headerColWidth: 50,
  headerRowHeight: 30,
  freezeRows: 3,
  freezeCols: 2,
  freezeOriginRow: 9,
  freezeOriginCol: 1,
  hiddenRows: new Set(),
  hiddenCols: new Set(),
  outlineRows: new Map(),
  outlineCols: new Map(),
  outlineRowGutter: 0,
  outlineColGutter: 0,
  hiddenSheets: new Set(),
  veryHiddenSheets: new Set(),
  sheetTabColors: new Map(),
};

const viewport: ViewportSlice = {
  rowStart: 12,
  rowCount: 5,
  colStart: 3,
  colCount: 4,
  widthPx: 0,
  zoom: 1,
};

describe('freeze band with an origin', () => {
  it('sizes the band from the origin, not from A1', () => {
    expect(frozenRowsHeight(layout, viewport)).toBe(60);
    expect(frozenColsWidth(layout, viewport)).toBe(200);
  });

  it('places the first band row/col at the data origin and the body after the band', () => {
    expect(cellRect(layout, viewport, 9, 1)).toMatchObject({ x: 50, y: 30 });
    expect(cellRect(layout, viewport, 12, 3)).toMatchObject({ x: 250, y: 90 });
  });

  it('renders the band then the body, skipping everything above/left of the origin', () => {
    expect(visibleRows(layout, viewport)).toEqual([9, 10, 11, 12, 13, 14, 15, 16]);
    expect(visibleCols(layout, viewport)).toEqual([1, 2, 3, 4, 5, 6]);
    expect(buildRowLayout(layout, viewport).frozenTotal).toBe(60);
    expect(buildColLayout(layout, viewport).visible[0]).toBe(1);
    expect(isRowVisible(layout, viewport, 8)).toBe(false);
    expect(isColVisible(layout, viewport, 0)).toBe(false);
    expect(isRowVisible(layout, viewport, 10)).toBe(true);
  });

  it('hit-tests the band and the body', () => {
    expect(hitTest(layout, viewport, 60, 40)).toEqual({ row: 9, col: 1 });
    expect(hitTest(layout, viewport, 260, 100)).toEqual({ row: 12, col: 3 });
  });

  it('splits a range across the band and the body', () => {
    const rects = rangeRects(layout, viewport, { r0: 10, r1: 13, c0: 2, c1: 3 });
    expect(rects).toHaveLength(4);
  });

  it('ignores the origin on an axis with nothing frozen', () => {
    const colsOnly = { ...layout, freezeRows: 0 };
    expect(visibleRows(colsOnly, { ...viewport, rowStart: 4 })[0]).toBe(4);
    expect(frozenRowsHeight(colsOnly, viewport)).toBe(0);
  });
});
