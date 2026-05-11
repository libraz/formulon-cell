import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  cellRect,
  cellRectIn,
  colLabel,
  colLeftEdge,
  colWidth,
  frozenColsWidth,
  frozenRowsHeight,
  hitTest,
  hitZone,
  isColVisible,
  isRowVisible,
  rangeRects,
  rowHeight,
  rowTopEdge,
  visibleCols,
  visibleRows,
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

describe('colLabel', () => {
  it('maps single-letter columns', () => {
    expect(colLabel(0)).toBe('A');
    expect(colLabel(1)).toBe('B');
    expect(colLabel(25)).toBe('Z');
  });

  it('maps two-letter columns', () => {
    expect(colLabel(26)).toBe('AA');
    expect(colLabel(27)).toBe('AB');
    expect(colLabel(51)).toBe('AZ');
    expect(colLabel(52)).toBe('BA');
    expect(colLabel(701)).toBe('ZZ');
  });

  it('maps three-letter columns', () => {
    expect(colLabel(702)).toBe('AAA');
    expect(colLabel(16383)).toBe('XFD'); // the spreadsheet last column
  });
});

describe('colWidth / rowHeight', () => {
  it('returns default when no override', () => {
    const layout = makeLayout();
    expect(colWidth(layout, 5)).toBe(100);
    expect(rowHeight(layout, 5)).toBe(20);
  });

  it('returns override when set', () => {
    const layout = makeLayout({
      colWidths: new Map([[3, 200]]),
      rowHeights: new Map([[7, 40]]),
    });
    expect(colWidth(layout, 3)).toBe(200);
    expect(rowHeight(layout, 7)).toBe(40);
  });

  it('returns 0 when hidden', () => {
    const layout = makeLayout({
      hiddenCols: new Set([2]),
      hiddenRows: new Set([4]),
    });
    expect(colWidth(layout, 2)).toBe(0);
    expect(rowHeight(layout, 4)).toBe(0);
  });

  it('scales visible dimensions by viewport zoom', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ zoom: 1.5 });
    expect(colWidth(layout, 5, viewport)).toBe(150);
    expect(rowHeight(layout, 5, viewport)).toBe(30);
  });
});

describe('frozenColsWidth / frozenRowsHeight', () => {
  it('returns 0 with no freeze', () => {
    const layout = makeLayout();
    expect(frozenColsWidth(layout)).toBe(0);
    expect(frozenRowsHeight(layout)).toBe(0);
  });

  it('sums frozen columns', () => {
    const layout = makeLayout({
      freezeCols: 3,
      colWidths: new Map([
        [0, 60],
        [1, 80],
      ]),
    });
    // 60 + 80 + 100 (default)
    expect(frozenColsWidth(layout)).toBe(240);
  });

  it('sums frozen rows', () => {
    const layout = makeLayout({
      freezeRows: 2,
      rowHeights: new Map([[1, 50]]),
    });
    // 20 + 50
    expect(frozenRowsHeight(layout)).toBe(70);
  });

  it('skips hidden frozen rows/cols', () => {
    const layout = makeLayout({
      freezeCols: 3,
      hiddenCols: new Set([1]),
    });
    // 100 + 0 + 100
    expect(frozenColsWidth(layout)).toBe(200);
  });
});

describe('cellRect', () => {
  it('positions a cell relative to headers', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Cell (2, 3) → x = headerCol + 3 cols = 50 + 300, y = headerRow + 2 rows = 30 + 40
    const r = cellRect(layout, viewport, 2, 3);
    expect(r.x).toBe(350);
    expect(r.y).toBe(70);
    expect(r.w).toBe(100);
    expect(r.h).toBe(20);
  });

  it('shifts non-frozen cells by colStart', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ colStart: 5 });
    // Body cell at col 7 sits 2 cols past viewport start.
    const r = cellRect(layout, viewport, 0, 7);
    expect(r.x).toBe(50 + 200); // header + 2 cols
  });

  it('keeps frozen cells anchored at the data origin', () => {
    const layout = makeLayout({ freezeCols: 2 });
    const viewport = makeViewport({ colStart: 10 });
    // Frozen col 0 stays at the very left of the data area.
    const frozen = cellRect(layout, viewport, 0, 0);
    expect(frozen.x).toBe(50);
    // Frozen col 1 sits one cell right of col 0.
    const frozen1 = cellRect(layout, viewport, 0, 1);
    expect(frozen1.x).toBe(150);
    // Body col 10 sits past the frozen band, with no scroll offset since
    // colStart is already at 10.
    const body = cellRect(layout, viewport, 0, 10);
    expect(body.x).toBe(50 + 200); // header + frozen band
  });

  it('applies zoom to cell positions and dimensions', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ zoom: 1.25 });
    const r = cellRect(layout, viewport, 2, 3);
    expect(r.x).toBe(50 + 375);
    expect(r.y).toBe(30 + 50);
    expect(r.w).toBe(125);
    expect(r.h).toBe(25);
  });
});

describe('visibleRows / visibleCols', () => {
  it('lists exactly the body window when no freeze and no hidden', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ rowStart: 5, rowCount: 3, colStart: 2, colCount: 4 });
    expect(visibleRows(layout, viewport)).toEqual([5, 6, 7]);
    expect(visibleCols(layout, viewport)).toEqual([2, 3, 4, 5]);
  });

  it('puts frozen rows/cols ahead of the body window', () => {
    const layout = makeLayout({ freezeRows: 2, freezeCols: 1 });
    const viewport = makeViewport({ rowStart: 5, rowCount: 2, colStart: 4, colCount: 2 });
    expect(visibleRows(layout, viewport)).toEqual([0, 1, 5, 6]);
    expect(visibleCols(layout, viewport)).toEqual([0, 4, 5]);
  });

  it('skips hidden indices in both bands', () => {
    const layout = makeLayout({
      freezeCols: 2,
      hiddenCols: new Set([1, 4]),
    });
    const viewport = makeViewport({ colStart: 3, colCount: 3 });
    expect(visibleCols(layout, viewport)).toEqual([0, 3, 5]);
  });

  it('clips body start to the freeze boundary', () => {
    // viewport.colStart < freezeCols — the body slice must start at freezeCols
    // so we never duplicate frozen cols.
    const layout = makeLayout({ freezeCols: 3 });
    const viewport = makeViewport({ colStart: 0, colCount: 5 });
    expect(visibleCols(layout, viewport)).toEqual([0, 1, 2, 3, 4]);
  });
});

describe('isRowVisible / isColVisible', () => {
  it('treats frozen indices as always visible (when not hidden)', () => {
    const layout = makeLayout({ freezeRows: 2, freezeCols: 1 });
    const viewport = makeViewport({ rowStart: 50, colStart: 50 });
    expect(isRowVisible(layout, viewport, 0)).toBe(true);
    expect(isColVisible(layout, viewport, 0)).toBe(true);
  });

  it('reports body indices inside the scrolled window', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ rowStart: 5, rowCount: 3, colStart: 2, colCount: 4 });
    expect(isRowVisible(layout, viewport, 6)).toBe(true);
    expect(isColVisible(layout, viewport, 5)).toBe(true);
    expect(isRowVisible(layout, viewport, 8)).toBe(false);
    expect(isColVisible(layout, viewport, 1)).toBe(false);
  });

  it('hidden indices are never visible, even in the frozen band', () => {
    const layout = makeLayout({ freezeCols: 3, hiddenCols: new Set([1]) });
    const viewport = makeViewport();
    expect(isColVisible(layout, viewport, 1)).toBe(false);
  });

  it('rejects negative indices', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    expect(isColVisible(layout, viewport, -1)).toBe(false);
    expect(isRowVisible(layout, viewport, -1)).toBe(false);
  });
});

describe('hitTest', () => {
  it('returns null inside a header strip', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // x < headerColWidth
    expect(hitTest(layout, viewport, 10, 200)).toBeNull();
    // y < headerRowHeight
    expect(hitTest(layout, viewport, 200, 10)).toBeNull();
  });

  it('resolves a body cell under the pointer', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // (col=2, row=3): x = 50 + 200 + half col, y = 30 + 60 + half row
    const got = hitTest(layout, viewport, 50 + 250, 30 + 70);
    expect(got).toEqual({ row: 3, col: 2 });
  });

  it('resolves a frozen cell when pointer is inside the frozen band', () => {
    const layout = makeLayout({ freezeCols: 2, freezeRows: 1 });
    const viewport = makeViewport({ colStart: 20, rowStart: 20 });
    // Pointer inside frozen col 1, frozen row 0.
    const got = hitTest(layout, viewport, 50 + 150, 30 + 10);
    expect(got).toEqual({ row: 0, col: 1 });
  });

  it('returns null past the last visible cell', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ colCount: 3 });
    // Past col 2's right edge.
    expect(hitTest(layout, viewport, 50 + 9999, 30 + 10)).toBeNull();
  });
});

describe('hitZone', () => {
  it('returns "corner" when both axes are inside header strips', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    expect(hitZone(layout, viewport, 5, 5)).toEqual({ kind: 'corner' });
  });

  it('returns "col-header" inside the column strip', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Past any resize zone, inside col 2.
    const z = hitZone(layout, viewport, 50 + 250, 5);
    expect(z).toEqual({ kind: 'col-header', col: 2 });
  });

  it('returns "row-header" inside the row strip', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Inside row 3, away from resize edges.
    const z = hitZone(layout, viewport, 5, 30 + 70);
    expect(z).toEqual({ kind: 'row-header', row: 3 });
  });

  it('detects col-resize at the right edge of a header cell', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Right edge of col 0 is x = 50 + 100 = 150.
    const z = hitZone(layout, viewport, 149, 5);
    expect(z).toEqual({ kind: 'col-resize', col: 0 });
  });

  it('detects col-resize at the left edge of a header cell (returns prior col)', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Left edge of col 1 is x = 150. Pointer just inside col 1 → resize prior col 0.
    const z = hitZone(layout, viewport, 151, 5);
    expect(z).toEqual({ kind: 'col-resize', col: 0 });
  });

  it('detects row-resize at the bottom edge of a row header', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Bottom edge of row 0 is y = 30 + 20 = 50.
    const z = hitZone(layout, viewport, 5, 49);
    expect(z).toEqual({ kind: 'row-resize', row: 0 });
  });

  it('returns a cell zone in the data area', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    const z = hitZone(layout, viewport, 50 + 250, 30 + 70);
    expect(z).toEqual({ kind: 'cell', row: 3, col: 2 });
  });

  it('returns "col-filter-btn" near the right edge of a header inside filterRange', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Col 0 right edge = 150. Resize slack = 4 → 146..150 is col-resize.
    // Filter btn is the next 14 px inboard: 132..146.
    const fr = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 };
    const z = hitZone(layout, viewport, 140, 5, fr);
    expect(z).toEqual({ kind: 'col-filter-btn', col: 0 });
  });

  it('still returns "col-resize" at the very right edge even when filterRange is set', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    const fr = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 2 };
    const z = hitZone(layout, viewport, 149, 5, fr);
    expect(z).toEqual({ kind: 'col-resize', col: 0 });
  });

  it('does not surface filter button outside filterRange columns', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    // Filter range covers col 0 only — col 1's right edge stays a regular header.
    const fr = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 };
    const z = hitZone(layout, viewport, 240, 5, fr); // col 1 mid
    expect(z).toEqual({ kind: 'col-header', col: 1 });
  });

  it('skips filter button when no filterRange is provided (legacy 4-arg call)', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    const z = hitZone(layout, viewport, 140, 5);
    expect(z).toEqual({ kind: 'col-header', col: 0 });
  });
});

describe('rangeRects', () => {
  it('returns one rect for a fully visible body range', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    const rects = rangeRects(layout, viewport, { r0: 1, r1: 2, c0: 1, c1: 2 });
    expect(rects).toHaveLength(1);
    const r = rects[0];
    if (!r) throw new Error('expected one rect');
    expect(r.x).toBe(50 + 100);
    expect(r.y).toBe(30 + 20);
    expect(r.w).toBe(200);
    expect(r.h).toBe(40);
  });

  it('returns empty when range is entirely past the viewport', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ rowCount: 2, colCount: 2 });
    const rects = rangeRects(layout, viewport, { r0: 50, r1: 60, c0: 50, c1: 60 });
    expect(rects).toEqual([]);
  });

  it('splits a range that crosses the freeze boundary into multiple rects', () => {
    // freezeRows=2, freezeCols=1; range covers both bands on each axis.
    const layout = makeLayout({ freezeRows: 2, freezeCols: 1 });
    const viewport = makeViewport({ rowStart: 5, colStart: 5 });
    const rects = rangeRects(layout, viewport, { r0: 0, r1: 6, c0: 0, c1: 6 });
    // Two row segments × two col segments = up to 4 rects.
    expect(rects.length).toBe(4);
  });

  it('clips ranges that start before viewport.rowStart but end inside it', () => {
    const layout = makeLayout();
    const viewport = makeViewport({ rowStart: 5, rowCount: 5 });
    const rects = rangeRects(layout, viewport, { r0: 3, r1: 7, c0: 0, c1: 1 });
    // Only the [5,7] portion is in the body window — single segment.
    expect(rects).toHaveLength(1);
  });
});

describe('colLeftEdge / rowTopEdge', () => {
  it('matches cellRect origin', () => {
    const layout = makeLayout();
    const viewport = makeViewport();
    const r = cellRect(layout, viewport, 4, 3);
    expect(colLeftEdge(layout, viewport, 3)).toBe(r.x);
    expect(rowTopEdge(layout, viewport, 4)).toBe(r.y);
  });
});

describe('buildColLayout / buildRowLayout', () => {
  it('orders frozen indices ahead of body slice', () => {
    const layout = makeLayout({ freezeCols: 2, freezeRows: 1 });
    const viewport = makeViewport({ rowStart: 5, rowCount: 2, colStart: 4, colCount: 2 });
    const cols = buildColLayout(layout, viewport);
    const rows = buildRowLayout(layout, viewport);
    expect(cols.visible).toEqual([0, 1, 4, 5]);
    expect(rows.visible).toEqual([0, 5, 6]);
  });

  it('records cumulative pixel positions and matches frozen totals', () => {
    const layout = makeLayout({
      freezeCols: 2,
      colWidths: new Map([
        [0, 60],
        [1, 80],
        [3, 50],
      ]),
    });
    const viewport = makeViewport({ colStart: 2, colCount: 4 });
    const cols = buildColLayout(layout, viewport);
    expect(cols.frozenTotal).toBe(140);
    expect(cols.positionAt.get(0)).toBe(0);
    expect(cols.positionAt.get(1)).toBe(60);
    // Body starts past the frozen band.
    expect(cols.positionAt.get(2)).toBe(140);
    expect(cols.positionAt.get(3)).toBe(140 + 100);
    // sizeAt mirrors colWidth.
    expect(cols.sizeAt.get(3)).toBe(50);
  });

  it('records zoomed cumulative positions and sizes', () => {
    const layout = makeLayout({
      freezeCols: 1,
      colWidths: new Map([[0, 80]]),
    });
    const viewport = makeViewport({ colStart: 1, colCount: 3, zoom: 1.5 });
    const cols = buildColLayout(layout, viewport);
    expect(cols.frozenTotal).toBe(120);
    expect(cols.positionAt.get(1)).toBe(120);
    expect(cols.positionAt.get(2)).toBe(270);
    expect(cols.sizeAt.get(2)).toBe(150);
  });

  it('skips hidden indices and adds no pixel space', () => {
    const layout = makeLayout({ hiddenCols: new Set([2]) });
    const viewport = makeViewport({ colStart: 0, colCount: 5 });
    const cols = buildColLayout(layout, viewport);
    expect(cols.visible).toEqual([0, 1, 3, 4]);
    // Col 3 follows col 1 directly (col 2 contributes 0 pixels).
    expect(cols.positionAt.get(3)).toBe(200);
  });

  it('matches frozenColsWidth / frozenRowsHeight in frozenTotal', () => {
    const layout = makeLayout({
      freezeCols: 3,
      freezeRows: 2,
      colWidths: new Map([[1, 70]]),
      rowHeights: new Map([[0, 35]]),
    });
    const viewport = makeViewport();
    const cols = buildColLayout(layout, viewport);
    const rows = buildRowLayout(layout, viewport);
    expect(cols.frozenTotal).toBe(frozenColsWidth(layout));
    expect(rows.frozenTotal).toBe(frozenRowsHeight(layout));
  });
});

describe('cellRectIn', () => {
  it('matches cellRect for visible cells across many shapes', () => {
    const layout = makeLayout({
      freezeCols: 2,
      freezeRows: 1,
      colWidths: new Map([
        [1, 60],
        [4, 130],
      ]),
      rowHeights: new Map([
        [0, 35],
        [5, 25],
      ]),
    });
    const viewport = makeViewport({ rowStart: 4, rowCount: 4, colStart: 3, colCount: 3 });
    const cols = buildColLayout(layout, viewport);
    const rows = buildRowLayout(layout, viewport);
    for (const r of rows.visible) {
      for (const c of cols.visible) {
        const a = cellRect(layout, viewport, r, c);
        const b = cellRectIn(layout, cols, rows, r, c);
        expect(b).toEqual(a);
      }
    }
  });
});
