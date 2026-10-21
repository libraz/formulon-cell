import { describe, expect, it } from 'vitest';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { ruleWidth } from '../../../../src/render/grid/chrome-context.js';
import { paintFreezeDividers, paintGridLines } from '../../../../src/render/grid/lines.js';
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

describe('paintFreezeDividers', () => {
  it('draws one CSS pixel in the freeze rule colour at the band edge', () => {
    const state = createSpreadsheetStore().getState();
    state.layout = { ...state.layout, freezeRows: 2, freezeCols: 1 };
    state.viewport = { rowStart: 2, rowCount: 2, colStart: 1, colCount: 2, zoom: 1, widthPx: 400 };
    const view = { ...state, layout: layoutForView(state) };
    const cols = buildColLayout(view.layout, view.viewport);
    const rows = buildRowLayout(view.layout, view.viewport);
    const strokes: string[] = [];
    const moves: [number, number][] = [];
    let lineWidth = 0;
    const ctx = {
      set strokeStyle(v: string) {
        strokes.push(v);
      },
      set lineWidth(v: number) {
        lineWidth = v;
      },
      beginPath: () => undefined,
      moveTo: (x: number, y: number) => moves.push([x, y]),
      lineTo: () => undefined,
      stroke: () => undefined,
    } as unknown as CanvasRenderingContext2D;

    paintFreezeDividers(
      { ctx, dpr: 2, cssWidth: 400, cssHeight: 200 },
      view,
      { freezeRule: '#9b9b9b', ruleStrong: '#cdc6b3' } as ResolvedTheme,
      cols,
      rows,
    );

    expect(strokes).toEqual(['#9b9b9b']);
    expect(lineWidth).toBe(1);
    // Header 20 + two 20px rows, header 32 + one 64px column, plus half the
    // rule width so the stroke covers a whole pixel.
    expect(moves).toEqual([
      [0, 60.5],
      [96.5, 0],
    ]);
  });
});
