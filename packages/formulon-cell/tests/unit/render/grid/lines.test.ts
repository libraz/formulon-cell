import { describe, expect, it } from 'vitest';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { ruleWidth } from '../../../../src/render/grid/chrome-context.js';
import { paintGridLines } from '../../../../src/render/grid/lines.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

describe('paintGridLines', () => {
  it('draws one CSS pixel at every display scale, snapped to device pixels', () => {
    expect(ruleWidth(1)).toBe(1);
    expect(ruleWidth(2)).toBe(1);
    expect(ruleWidth(1.5)).toBeCloseTo(2 / 1.5);

    const state = createSpreadsheetStore().getState();
    state.viewport = { rowStart: 0, rowCount: 2, colStart: 0, colCount: 2, zoom: 1, widthPx: 200 };
    const view = { ...state, layout: layoutForView(state) };
    const cols = buildColLayout(view.layout, view.viewport);
    const rows = buildRowLayout(view.layout, view.viewport);
    const moves: number[] = [];
    let lineWidth = 0;
    const ctx = {
      set strokeStyle(_v: string) {},
      set lineWidth(v: number) {
        lineWidth = v;
      },
      beginPath: () => undefined,
      moveTo: (x: number) => moves.push(x),
      lineTo: () => undefined,
      stroke: () => undefined,
    } as unknown as CanvasRenderingContext2D;

    paintGridLines(
      { ctx, dpr: 2, cssWidth: 200, cssHeight: 100 },
      view,
      { rule: '#d4d4d4' } as ResolvedTheme,
      cols,
      rows,
    );

    expect(lineWidth).toBe(1);
    // A rule sits on the cell's leading edge and covers one whole CSS pixel.
    expect(moves.every((x) => x === 0 || x % 1 === 0.5)).toBe(true);
  });
});
