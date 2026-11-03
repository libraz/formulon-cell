import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  gridOriginX,
  gridOriginY,
  layoutForView,
} from '../../../../src/render/geometry.js';
import { paintHeaders } from '../../../../src/render/grid/headers.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = { accent: '#217346', headerBg: '#fff', fontUi: 'sans-serif', textHeader: 11 };

describe('paintHeaders hidden span marks', () => {
  it('draws a 6 device px accent bar across the strip at each hidden boundary', () => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenCols: new Set([2]), hiddenRows: new Set([2, 3]) },
      viewport: { ...s.viewport, rowStart: 0, rowCount: 8, colStart: 0, colCount: 5 },
    }));
    const state = store.getState();
    const layout = layoutForView(state);
    const cols = buildColLayout(layout, state.viewport);
    const rows = buildRowLayout(layout, state.viewport);
    const rects: Array<{ style: string; rect: number[] }> = [];
    let style = '';
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'fillStyle'
            ? style
            : key === 'fillRect'
              ? (...rect: number[]) => rects.push({ style, rect })
              : key === 'measureText'
                ? () => ({ width: 0 })
                : () => undefined,
        set: (_t, key, v) => {
          if (key === 'fillStyle') style = String(v);
          return true;
        },
      },
    ) as unknown as CanvasRenderingContext2D;

    paintHeaders(
      { ctx, dpr: 2, cssWidth: 600, cssHeight: 400 },
      { ...state, layout },
      theme as unknown as ResolvedTheme,
      cols,
      rows,
    );

    const bars = rects.filter((r) => r.style === '#217346').map((r) => r.rect);
    const colBoundary = cellRectIn(layout, cols, rows, 0, 1);
    const rowBoundary = cellRectIn(layout, cols, rows, 1, 0);
    const x = colBoundary.x + colBoundary.w - 2;
    const y = rowBoundary.y + rowBoundary.h - 2;
    const oy = gridOriginY(layout);
    const ox = gridOriginX(layout);
    expect(bars).toContainEqual([x, oy - layout.headerRowHeight, 3, layout.headerRowHeight]);
    expect(bars).toContainEqual([ox - layout.headerColWidth, y, layout.headerColWidth, 3]);
  });
});

describe('headerLabelSize', () => {
  it('follows a Normal font larger than the default, else keeps the theme size', async () => {
    const { headerLabelSize } = await import('../../../../src/render/grid.js');
    expect(headerLabelSize(20, 11.5)).toBe(20);
    expect(headerLabelSize(11, 11.5)).toBe(11.5);
    expect(headerLabelSize(undefined, 11.5)).toBe(11.5);
  });
});
