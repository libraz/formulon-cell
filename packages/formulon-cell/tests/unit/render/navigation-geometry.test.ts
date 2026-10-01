import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  colX,
  frozenColsWidth,
  frozenRowsHeight,
  hitTest,
  rowY,
  type ViewLayout,
} from '../../../src/render/geometry.js';
import type { ViewportSlice } from '../../../src/store/store.js';

const layout: ViewLayout = {
  rtl: false,
  viewWidth: 500,
  colWidths: new Map(),
  rowHeights: new Map(),
  defaultColWidth: 80,
  defaultRowHeight: 20,
  headerColWidth: 30,
  headerRowHeight: 20,
  freezeRows: 0,
  freezeCols: 0,
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
  rowStart: 10,
  rowCount: 40,
  colStart: 5,
  colCount: 20,
  zoom: 1,
  widthPx: 500,
  navigationRange: { sheet: 0, r0: 10, c0: 5, r1: 11, c1: 6 },
};

describe('bounded navigation geometry', () => {
  it('paints only rows and columns inside the finite navigation rectangle', () => {
    expect(buildRowLayout(layout, viewport).visible).toEqual([10, 11]);
    expect(buildColLayout(layout, viewport).visible).toEqual([5, 6]);
  });

  it('does not hit-test phantom cells beyond the finite rectangle', () => {
    // Header (30 × 20), then two 80px columns and two 20px rows.
    expect(hitTest(layout, viewport, 31, 21)).toEqual({ row: 10, col: 5 });
    expect(hitTest(layout, viewport, 150, 21)).toEqual({ row: 10, col: 6 });
    expect(hitTest(layout, viewport, 270, 21)).toBeNull();
    expect(hitTest(layout, viewport, 31, 65)).toBeNull();
  });

  it('keeps a nonzero bounded region aligned when it overlaps frozen rows and columns', () => {
    const frozenLayout: ViewLayout = { ...layout, freezeRows: 3, freezeCols: 3 };
    const frozenViewport: ViewportSlice = {
      ...viewport,
      rowStart: 1,
      colStart: 1,
      navigationRange: { sheet: 0, r0: 1, c0: 1, r1: 4, c1: 4 },
    };

    expect(buildRowLayout(frozenLayout, frozenViewport).visible).toEqual([1, 2, 3, 4]);
    expect(buildColLayout(frozenLayout, frozenViewport).visible).toEqual([1, 2, 3, 4]);
    expect(frozenRowsHeight(frozenLayout, frozenViewport)).toBe(40);
    expect(frozenColsWidth(frozenLayout, frozenViewport)).toBe(160);
    expect(rowY(frozenLayout, frozenViewport, 1)).toBe(0);
    expect(rowY(frozenLayout, frozenViewport, 2)).toBe(20);
    expect(rowY(frozenLayout, frozenViewport, 3)).toBe(40);
    expect(colX(frozenLayout, frozenViewport, 1)).toBe(0);
    expect(colX(frozenLayout, frozenViewport, 2)).toBe(80);
    expect(colX(frozenLayout, frozenViewport, 3)).toBe(160);

    expect(hitTest(frozenLayout, frozenViewport, 31, 21)).toEqual({ row: 1, col: 1 });
    expect(hitTest(frozenLayout, frozenViewport, 111, 21)).toEqual({ row: 1, col: 2 });
    expect(hitTest(frozenLayout, frozenViewport, 191, 61)).toEqual({ row: 3, col: 3 });
  });
});
