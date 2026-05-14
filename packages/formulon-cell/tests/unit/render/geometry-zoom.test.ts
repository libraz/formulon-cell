import { describe, expect, it } from 'vitest';

import {
  cellRect,
  colWidth,
  frozenColsWidth,
  frozenRowsHeight,
  hitTest,
  rowHeight,
} from '../../../src/render/geometry.js';
import type { LayoutSlice, ViewportSlice } from '../../../src/store/store.js';

function makeLayout(over: Partial<LayoutSlice> = {}): LayoutSlice {
  return {
    colWidths: new Map(),
    rowHeights: new Map(),
    defaultColWidth: 100,
    defaultRowHeight: 20,
    headerColWidth: 50,
    headerRowHeight: 30,
    freezeRows: 0,
    freezeCols: 0,
    hiddenRows: new Set(),
    hiddenCols: new Set(),
    outlineRows: new Map(),
    outlineCols: new Map(),
    outlineRowGutter: 0,
    outlineColGutter: 0,
    hiddenSheets: new Set(),
    ...over,
  };
}

function makeViewport(over: Partial<ViewportSlice> = {}): ViewportSlice {
  return {
    rowStart: 0,
    rowCount: 10,
    colStart: 0,
    colCount: 6,
    zoom: 1,
    ...over,
  };
}

describe('render/geometry — zoom scaling', () => {
  describe.each([
    [0.5, 50, 10],
    [1.0, 100, 20],
    [1.5, 150, 30],
    [2.0, 200, 40],
  ])('zoom %s', (zoom, expectedW, expectedH) => {
    const layout = makeLayout();
    const viewport = makeViewport({ zoom });

    it(`scales colWidth to ${expectedW}px`, () => {
      expect(colWidth(layout, 0, viewport)).toBe(expectedW);
    });

    it(`scales rowHeight to ${expectedH}px`, () => {
      expect(rowHeight(layout, 0, viewport)).toBe(expectedH);
    });

    it(`places cellRect (0,0) past the header at the correct device origin`, () => {
      const r = cellRect(layout, viewport, 0, 0);
      expect(r.x).toBe(50); // headerColWidth — NOT zoomed; header is chrome
      expect(r.y).toBe(30);
      expect(r.w).toBe(expectedW);
      expect(r.h).toBe(expectedH);
    });
  });

  it('returns zero for hidden rows / cols even when zoom is non-1', () => {
    const layout = makeLayout({ hiddenCols: new Set([0]), hiddenRows: new Set([0]) });
    const viewport = makeViewport({ zoom: 2 });
    expect(colWidth(layout, 0, viewport)).toBe(0);
    expect(rowHeight(layout, 0, viewport)).toBe(0);
  });

  it('honours per-cell colWidths overrides with zoom', () => {
    const layout = makeLayout({ colWidths: new Map([[0, 120]]) });
    const viewport = makeViewport({ zoom: 1.5 });
    expect(colWidth(layout, 0, viewport)).toBe(180);
    expect(colWidth(layout, 1, viewport)).toBe(150); // default 100 × 1.5
  });

  it('scales frozen column/row bands by zoom', () => {
    const layout = makeLayout({ freezeCols: 2, freezeRows: 1 });
    const viewport = makeViewport({ zoom: 1.5 });
    // 2 cols × 100 × 1.5
    expect(frozenColsWidth(layout, viewport)).toBe(300);
    // 1 row × 20 × 1.5
    expect(frozenRowsHeight(layout, viewport)).toBe(30);
  });
});

describe('render/geometry — hitTest under zoom', () => {
  it('maps device x into the zoomed grid columns', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ zoom: 2 });
    // 50px header + 200 (col 0 zoomed) → x=251 should land in col 1
    expect(hitTest(layout, viewport, 60, 35)?.col).toBe(0);
    expect(hitTest(layout, viewport, 251, 35)?.col).toBe(1);
    expect(hitTest(layout, viewport, 451, 35)?.col).toBe(2);
  });

  it('returns null inside the header strip', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    expect(hitTest(layout, viewport, 10, 10)).toBeNull();
    expect(hitTest(layout, viewport, 49, 50)).toBeNull(); // x < headerColWidth
  });

  it('resolves clicks inside the frozen band to a frozen column', () => {
    const layout = makeLayout({ freezeCols: 2 });
    const viewport = makeViewport({ zoom: 1, colStart: 5 }); // body starts at col 5
    // header (50) + 100 (col 0) = 150 → col 1
    expect(hitTest(layout, viewport, 175, 50)?.col).toBe(1);
  });
});
