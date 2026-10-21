import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { pageLayoutGaps, resetPaginationCache } from '../../../src/commands/pagination.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  buildColLayout,
  buildRowLayout,
  cellRect,
  cellRectIn,
  colWidth,
  hitTest,
  rowHeight,
  type ViewLayout,
} from '../../../src/render/geometry.js';
import { GridRenderer } from '../../../src/render/grid.js';
import {
  createSpreadsheetStore,
  type SpreadsheetStore,
  type State,
  type ViewportSlice,
} from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

// Mac new-book defaults in canonical px: 75 pt columns, 20 pt rows.
const COL = 100;
const ROW = 80 / 3;
const HEADER_W = 26;
const HEADER_H = 20;

function makeLayout(over: Partial<ViewLayout> = {}): ViewLayout {
  return {
    rtl: false,
    viewWidth: 0,
    colWidths: new Map(),
    rowHeights: new Map(),
    defaultColWidth: COL,
    defaultRowHeight: ROW,
    headerColWidth: HEADER_W,
    headerRowHeight: HEADER_H,
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
    ...over,
  };
}

function makeViewport(over: Partial<ViewportSlice> = {}): ViewportSlice {
  return {
    rowStart: 0,
    rowCount: 10,
    colStart: 0,
    colCount: 6,
    widthPx: 0,
    zoom: 1,
    ...over,
  };
}

describe.each([
  [1, 75, 20],
  [1.5, 112.5, 30],
])('render geometry at screenScale 0.75, zoom %s', (zoom, w, h) => {
  const layout = makeLayout({ screenScale: 0.75 });
  const viewport = makeViewport({ zoom });

  it('sizes columns and rows in screen px', () => {
    expect(colWidth(layout, 0, viewport)).toBeCloseTo(w, 9);
    expect(rowHeight(layout, 0, viewport)).toBeCloseTo(h, 9);
  });

  it('places cellRect past unscaled headers', () => {
    const r = cellRect(layout, viewport, 1, 2);
    expect(r.x).toBeCloseTo(HEADER_W + 2 * w, 9);
    expect(r.y).toBeCloseTo(HEADER_H + h, 9);
    expect(r.w).toBeCloseTo(w, 9);
    expect(r.h).toBeCloseTo(h, 9);
  });

  it('hit-tests screen px against the scaled grid', () => {
    expect(hitTest(layout, viewport, HEADER_W + w - 1, HEADER_H + h - 1)).toEqual({
      row: 0,
      col: 0,
    });
    expect(hitTest(layout, viewport, HEADER_W + w + 1, HEADER_H + h + 1)).toEqual({
      row: 1,
      col: 1,
    });
  });

  it('agrees with cellRectIn for laid-out cells', () => {
    const cols = buildColLayout(layout, viewport);
    const rows = buildRowLayout(layout, viewport);
    expect(cellRectIn(layout, cols, rows, 1, 2)).toEqual(cellRect(layout, viewport, 1, 2));
  });
});

describe('cellRectIn fallback without a viewport', () => {
  it('scales the nominal size by screenScale', () => {
    const layout = makeLayout({ screenScale: 0.75 });
    const empty = buildColLayout(layout, makeViewport({ colCount: 0, rowCount: 0 }));
    const r = cellRectIn(layout, empty, empty, 50, 50);
    expect(r.w).toBeCloseTo(75, 9);
    expect(r.h).toBeCloseTo(20, 9);
  });

  it('keeps canonical px when screenScale is absent', () => {
    const layout = makeLayout();
    const empty = buildColLayout(layout, makeViewport({ colCount: 0, rowCount: 0 }));
    const r = cellRectIn(layout, empty, empty, 50, 50);
    expect(r.w).toBe(COL);
    expect(r.h).toBe(ROW);
  });
});

describe('absent screenScale', () => {
  it.each([1, 1.5])('matches the zoom-only geometry at zoom %s', (zoom) => {
    const layout = makeLayout();
    const viewport = makeViewport({ zoom });
    expect(colWidth(layout, 0, viewport)).toBe(COL * zoom);
    expect(rowHeight(layout, 0, viewport)).toBe(ROW * zoom);
    expect(cellRect(layout, viewport, 1, 2)).toEqual({
      x: HEADER_W + 2 * COL * zoom,
      y: HEADER_H + ROW * zoom,
      w: COL * zoom,
      h: ROW * zoom,
    });
    expect(hitTest(layout, viewport, HEADER_W + COL * zoom + 1, HEADER_H + 1)).toEqual({
      row: 0,
      col: 1,
    });
  });
});

describe('visible row/col counts', () => {
  const theme = { textCell: 12, fontUi: 'Arial' } as unknown as ResolvedTheme;

  function viewportCounts(state: State): { rows: number; cols: number } {
    const counts = { rows: 0, cols: 0 };
    const host = {
      getBoundingClientRect: () => ({ width: HEADER_W + 610, height: HEADER_H + 250 }),
    } as unknown as HTMLElement;
    const canvas = {
      getContext: () => ({}) as CanvasRenderingContext2D,
      style: {},
    } as unknown as HTMLCanvasElement;
    const renderer = new GridRenderer({
      host,
      canvas,
      getState: () => state,
      getTheme: () => theme,
      onViewportSize: (rows, cols) => {
        counts.rows = rows;
        counts.cols = cols;
      },
    });
    renderer.resize();
    renderer.dispose();
    return counts;
  }

  function stateWith(screenScale: number | undefined, zoom: number): State {
    const state = createSpreadsheetStore().getState();
    return {
      ...state,
      layout: {
        ...state.layout,
        defaultColWidth: COL,
        defaultRowHeight: ROW,
        headerColWidth: HEADER_W,
        headerRowHeight: HEADER_H,
        ...(screenScale === undefined ? {} : { screenScale }),
      },
      viewport: { ...state.viewport, zoom },
    };
  }

  it.each([
    // Body 610 × 250, plus one trailing partial per axis: 250 / 20, 610 / 75.
    [0.75, 1, 14, 10],
    // 250 / 30, 610 / 112.5.
    [0.75, 1.5, 10, 7],
    // 250 / 26.67, 610 / 100.
    [undefined, 1, 11, 8],
  ])('screenScale %s, zoom %s fits %s rows × %s cols', (screenScale, zoom, rows, cols) => {
    expect(viewportCounts(stateWith(screenScale, zoom))).toEqual({ rows, cols });
  });
});

describe('Page Layout gutters', () => {
  let store: SpreadsheetStore;
  beforeEach(() => {
    resetPaginationCache();
    store = createSpreadsheetStore();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey({ sheet: 0, row: 200, col: 0 }), {
        value: { kind: 'number', value: 1 },
        formula: null,
      });
      return { ...s, data: { ...s.data, cells } };
    });
  });
  afterEach(() => resetPaginationCache());

  it('scale with screenScale like the cells they sit between', () => {
    const base = pageLayoutGaps(store.getState(), 0).pageGapRows.get(0) ?? 0;
    store.setState((s) => ({ ...s, layout: { ...s.layout, screenScale: 0.75 } }));
    resetPaginationCache();
    const scaled = pageLayoutGaps(store.getState(), 0).pageGapRows.get(0) ?? 0;
    expect(base).toBeGreaterThan(0);
    expect(scaled).toBeCloseTo(base * 0.75, 5);
  });
});
