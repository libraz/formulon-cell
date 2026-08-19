import { beforeEach, describe, expect, it } from 'vitest';
import { resetPaginationCache } from '../../../src/commands/pagination.js';
import { setWorkbookView } from '../../../src/commands/view.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  buildColLayout,
  buildRowLayout,
  cellRect,
  gridOriginX,
  gridOriginY,
  hitZone,
  layoutForView,
  RULER_BAND,
} from '../../../src/render/geometry.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const seed = (store: SpreadsheetStore, row: number, col: number): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), {
      value: { kind: 'number', value: 1 },
      formula: null,
    });
    return { ...s, data: { ...s.data, cells } };
  });
};

const wide = (store: SpreadsheetStore): void => {
  store.setState((s) => ({
    ...s,
    viewport: { ...s.viewport, rowCount: 120, colCount: 40, widthPx: 1200 },
  }));
};

describe('Page Layout geometry', () => {
  beforeEach(() => resetPaginationCache());

  it('leaves the axis untouched in Normal view', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    wide(store);
    const layout = layoutForView(store.getState());

    expect(layout.pageGapRows).toBeUndefined();
    expect(layout.pageGapCols).toBeUndefined();
    expect(layout.rulerRowHeight).toBeUndefined();
  });

  it('opens page gutters and reserves the ruler bands in Page Layout view', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    wide(store);
    const normalOrigin = {
      x: gridOriginX(layoutForView(store.getState())),
      y: gridOriginY(layoutForView(store.getState())),
    };

    setWorkbookView(store, 'pageLayout');
    const layout = layoutForView(store.getState());

    expect(layout.rulerRowHeight).toBe(RULER_BAND);
    expect(layout.rulerColWidth).toBe(RULER_BAND);
    expect(gridOriginX(layout)).toBe(normalOrigin.x + RULER_BAND);
    expect(gridOriginY(layout)).toBe(normalOrigin.y + RULER_BAND);
    // The first page's leading margin sits before row 0 / column 0.
    expect(layout.pageGapRows?.get(0)).toBeGreaterThan(0);
    expect(layout.pageGapCols?.get(0)).toBeGreaterThan(0);
  });

  it('pushes every cell past the page it opens', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    wide(store);
    const before = cellRect(layoutForView(store.getState()), store.getState().viewport, 0, 0);

    setWorkbookView(store, 'pageLayout');
    const layout = layoutForView(store.getState());
    const after = cellRect(layout, store.getState().viewport, 0, 0);

    expect(after.x).toBeGreaterThan(before.x);
    expect(after.y).toBeGreaterThan(before.y);
    expect(after.w).toBe(before.w);
  });

  it('reports the gutter alongside the cells it precedes', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    wide(store);
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const layout = layoutForView(state);
    const rows = buildRowLayout(layout, state.viewport);
    const cols = buildColLayout(layout, state.viewport);

    expect(rows.gapAt.get(0)).toBe(layout.pageGapRows?.get(0));
    expect(cols.gapAt.get(0)).toBe(layout.pageGapCols?.get(0));
    // `positionAt` already sits past the gutter.
    expect(rows.positionAt.get(0)).toBe(layout.pageGapRows?.get(0));
  });

  it('resolves no cell inside a page gutter', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    wide(store);
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const layout = layoutForView(state);
    const gap = layout.pageGapRows?.get(0) ?? 0;
    const lead = layout.pageGapCols?.get(0) ?? 0;
    expect(gap).toBeGreaterThan(2);

    const inGutter = hitZone(
      layout,
      state.viewport,
      gridOriginX(layout) + lead + 5,
      gridOriginY(layout) + 1,
    );
    expect(inGutter).toBeNull();

    const inCell = hitZone(
      layout,
      state.viewport,
      gridOriginX(layout) + lead + 5,
      gridOriginY(layout) + gap + 2,
    );
    expect(inCell).toEqual({ kind: 'cell', row: 0, col: 0 });
  });

  it('claims no row or column inside a ruler band', () => {
    const store = createSpreadsheetStore();
    seed(store, 200, 0);
    wide(store);
    setWorkbookView(store, 'pageLayout');
    const state = store.getState();
    const layout = layoutForView(state);

    expect(hitZone(layout, state.viewport, 200, RULER_BAND - 2)).toBeNull();
    expect(hitZone(layout, state.viewport, RULER_BAND - 2, 200)).toBeNull();
    // Just inboard of the band the header rails take over again.
    expect(hitZone(layout, state.viewport, 200, RULER_BAND + 2)?.kind).toBe('col-header');
  });
});
