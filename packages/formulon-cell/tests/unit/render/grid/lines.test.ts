import { describe, expect, it } from 'vitest';
import {
  buildColLayout,
  buildRowLayout,
  cellRectIn,
  layoutForView,
} from '../../../../src/render/geometry.js';
import { ruleWidth } from '../../../../src/render/grid/chrome-context.js';
import {
  gridLinesShown,
  paintFreezeDividers,
  paintGridLines,
} from '../../../../src/render/grid/lines.js';
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
    // Each cell owns the rule on its trailing edge: [E - 1px, E) centres half a CSS pixel before E.
    expect(moves.length).toBeGreaterThan(0);
    expect(moves.every((x) => x % 1 === 0.5 || x === 0)).toBe(true);
  });
});

describe('paintGridLines inside a PivotTable', () => {
  it('drops the rules the report cells own, keeping its trailing edge', () => {
    const state = createSpreadsheetStore().getState();
    state.viewport = { rowStart: 0, rowCount: 8, colStart: 0, colCount: 4, zoom: 1, widthPx: 400 };
    const view = { ...state, layout: layoutForView(state) };
    const cols = buildColLayout(view.layout, view.viewport);
    const rows = buildRowLayout(view.layout, view.viewport);
    const lines: [number, number, number, number][] = [];
    let from: [number, number] = [0, 0];
    const ctx = {
      set strokeStyle(_v: string) {},
      set lineWidth(_v: number) {},
      beginPath: () => undefined,
      moveTo: (x: number, y: number) => {
        from = [x, y];
      },
      lineTo: (x: number, y: number) => lines.push([...from, x, y]),
      stroke: () => undefined,
    } as unknown as CanvasRenderingContext2D;
    const pivot = { sheet: 0, r0: 2, c0: 0, r1: 4, c1: 1 };
    paintGridLines(
      { ctx, dpr: 1, cssWidth: 400, cssHeight: 300 },
      view,
      { rule: '#d4d4d4' } as ResolvedTheme,
      cols,
      rows,
      [pivot],
    );
    const at = (r: number, c: number) => cellRectIn(view.layout, cols, rows, r, c);
    const left = at(0, 0).x;
    const inner = at(0, 1).x;
    const right = at(0, 2).x;
    const top = at(2, 0).y;
    const bottom = at(5, 0).y;
    const vertical = (x: number) => lines.filter((l) => l[0] === x - 0.5 && l[2] === x - 0.5);
    const horizontal = (y: number) => lines.filter((l) => l[1] === y - 0.5 && l[3] === y - 0.5);
    // Between the report's columns: broken over the report's rows.
    expect(vertical(inner).map((l) => [l[1], l[3]])).toEqual([
      [0, top],
      [bottom, 300],
    ]);
    // The report's trailing edge and the rule before it run through.
    expect(vertical(right).map((l) => [l[1], l[3]])).toEqual([[0, 300]]);
    expect(vertical(left).map((l) => [l[1], l[3]])).toEqual([[0, 300]]);
    // Under each report row, the bottom one included: broken across the report.
    for (const r of [3, 4, 5]) {
      expect(horizontal(at(r, 0).y).map((l) => [l[0], l[2]])).toEqual([
        [0, left],
        [right, 400],
      ]);
    }
    // The rule above the report belongs to the row before it.
    expect(horizontal(top).map((l) => [l[0], l[2]])).toEqual([[0, 400]]);
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
    // Header 20 + two 20px rows, header 32 + one 64px column, minus half the
    // rule width so the stroke covers the pixel just before the boundary.
    expect(moves).toEqual([
      [0, 59.5],
      [95.5, 0],
    ]);
  });
});

describe('gridLinesShown', () => {
  // Excel for Mac draws no gridlines at 33% and 25% zoom, and draws them at 40%.
  const at = (zoom: number, screenScale: number, showGridLines = true) => {
    const state = createSpreadsheetStore().getState();
    return {
      ...state,
      layout: { ...state.layout, screenScale },
      viewport: { ...state.viewport, zoom },
      ui: { ...state.ui, showGridLines },
    };
  };

  it('hides Mac gridlines below 40% zoom', () => {
    expect(gridLinesShown(at(1, 0.75))).toBe(true);
    expect(gridLinesShown(at(0.4, 0.75))).toBe(true);
    expect(gridLinesShown(at(0.33, 0.75))).toBe(false);
    expect(gridLinesShown(at(0.25, 0.75))).toBe(false);
  });

  it('keeps the default profile and the sheet toggle as they were', () => {
    expect(gridLinesShown(at(0.25, 1))).toBe(true);
    expect(gridLinesShown(at(1, 0.75, false))).toBe(false);
  });
});
