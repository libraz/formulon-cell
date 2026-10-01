import { afterEach, describe, expect, it } from 'vitest';

import {
  buildColLayout,
  buildRowLayout,
  layoutForView,
  rangeRects,
} from '../../../../src/render/geometry.js';
import {
  getFillHandleRect,
  getValidationChevron,
  setFillHandleRect,
  setValidationChevron,
} from '../../../../src/render/grid/hit-state.js';
import { GridRenderer } from '../../../../src/render/grid.js';
import { createSpreadsheetStore, mutators, type State } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = {
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  bg: '#ffffff',
  bgElev: '#f5f5f5',
  bgRail: '#eeeeee',
  fg: '#1f1f1f',
  rule: '#cccccc',
  fontUi: 'Arial',
  fontMono: 'monospace',
  textCell: 12,
} as unknown as ResolvedTheme;

interface CanvasSpy {
  ctx: CanvasRenderingContext2D;
  fills: { style: string; rect: [number, number, number, number] }[];
  strokeRects: [number, number, number, number][];
  clipRects: [number, number, number, number][];
  textDraws: { text: string; x: number; y: number }[];
  strokes: number;
  dashes: number[][];
}

function canvasSpy(): CanvasSpy {
  const fills: CanvasSpy['fills'] = [];
  const strokeRects: CanvasSpy['strokeRects'] = [];
  const clipRects: CanvasSpy['clipRects'] = [];
  const textDraws: CanvasSpy['textDraws'] = [];
  const dashes: number[][] = [];
  let fillStyle = '';
  const spy = {
    ctx: null as unknown as CanvasRenderingContext2D,
    fills,
    strokeRects,
    clipRects,
    textDraws,
    strokes: 0,
    dashes,
  };
  const ctx = {
    get fillStyle(): string {
      return fillStyle;
    },
    set fillStyle(value: string) {
      fillStyle = value;
    },
    strokeStyle: '',
    lineWidth: 1,
    lineDashOffset: 0,
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'inherit',
    globalAlpha: 1,
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    closePath(): void {},
    moveTo(): void {},
    lineTo(): void {},
    rect(x: number, y: number, w: number, h: number): void {
      clipRects.push([x, y, w, h]);
    },
    clip(): void {},
    translate(): void {},
    rotate(): void {},
    ellipse(): void {},
    arc(): void {},
    fill(): void {},
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style: fillStyle, rect: [x, y, w, h] });
    },
    stroke(): void {
      spy.strokes += 1;
    },
    strokeRect(x: number, y: number, w: number, h: number): void {
      strokeRects.push([x, y, w, h]);
    },
    setLineDash(value: number[]): void {
      dashes.push([...value]);
    },
    fillText(text: string, x: number, y: number): void {
      textDraws.push({ text, x, y });
    },
    measureText(): TextMetrics {
      return { width: 0 } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
  spy.ctx = ctx;
  return spy;
}

function makeRenderer(state: State, spy: CanvasSpy): GridRenderer {
  const host = {
    getBoundingClientRect: () => ({ width: 640, height: 480 }),
  } as unknown as HTMLElement;
  const canvas = {
    getContext: () => spy.ctx,
    style: {},
  } as unknown as HTMLCanvasElement;
  return new GridRenderer({
    host,
    canvas,
    getState: () => state,
    getTheme: () => theme,
  });
}

function viewParts(state: State) {
  const view = { ...state, layout: layoutForView(state) };
  return {
    view,
    cols: buildColLayout(view.layout, view.viewport),
    rows: buildRowLayout(view.layout, view.viewport),
  };
}

function paintActive(renderer: GridRenderer, state: State): void {
  const { view, cols, rows } = viewParts(state);
  (
    renderer as unknown as {
      paintActive(state: State, theme: ResolvedTheme, cols: unknown, rows: unknown): void;
    }
  ).paintActive(view, theme, cols, rows);
}

function paintCells(renderer: GridRenderer, state: State): void {
  const { view, cols, rows } = viewParts(state);
  (
    renderer as unknown as {
      paintCells(state: State, theme: ResolvedTheme, cols: unknown, rows: unknown): void;
    }
  ).paintCells(view, theme, cols, rows);
}

function paintBorders(renderer: GridRenderer, state: State): void {
  const { view, cols, rows } = viewParts(state);
  (
    renderer as unknown as {
      paintBorders(state: State, theme: ResolvedTheme, cols: unknown, rows: unknown): void;
    }
  ).paintBorders(view, theme, cols, rows);
}

afterEach(() => {
  setFillHandleRect(null);
  setValidationChevron(null);
});

describe('GridRenderer merged selection parity', () => {
  it('uses the full merge for the active outline, validation chevron, and fill handle', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      {
        validation: { kind: 'list', source: ['one', 'two'] },
      },
    );
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintActive(renderer, state);

    const { view } = viewParts(state);
    const [mergeRect] = rangeRects(view.layout, view.viewport, state.selection.range);
    if (!mergeRect) throw new Error('expected the merged range to be visible');
    expect(spy.strokeRects[0]).toEqual([
      mergeRect.x + 1,
      mergeRect.y + 1,
      mergeRect.w - 2,
      mergeRect.h - 2,
    ]);
    // The selection outline is the same rectangle and must not be painted a
    // second time over the active merge outline. The chevron adds one stroke.
    expect(spy.strokeRects).toHaveLength(2);
    expect(getValidationChevron()).toMatchObject({ row: 1, col: 1 });
    expect(getValidationChevron()?.rect.x).toBe(mergeRect.x + mergeRect.w - 18);
    expect(getFillHandleRect()).toMatchObject({
      x: mergeRect.x + mergeRect.w - 4,
      y: mergeRect.y + mergeRect.h - 4,
    });
  });

  it('clears an inactive unfilled merge before content painting', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    mutators.setActive(store, { sheet: 0, row: 3, col: 3 });
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintCells(renderer, state);

    const { view } = viewParts(state);
    const [mergeRect] = rangeRects(view.layout, view.viewport, {
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    if (!mergeRect) throw new Error('expected the merged range to be visible');
    expect(spy.fills).toContainEqual({
      style: theme.bg,
      rect: [mergeRect.x, mergeRect.y, mergeRect.w, mergeRect.h],
    });
  });

  it('clips merged text to the union of contiguous freeze quadrants', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 3, 3);
    const merge = { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 4 } as const;
    mutators.setCell(
      store,
      { sheet: 0, row: merge.r0, col: merge.c0 },
      {
        kind: 'text',
        value: 'merged',
      },
    );
    mutators.mergeRange(store, merge);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: merge.r0, col: merge.c0 },
      { align: 'center', vAlign: 'middle' },
    );
    mutators.setActive(store, { sheet: 0, row: merge.r1, col: merge.c1 });
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintCells(renderer, state);

    const { view } = viewParts(state);
    const visibleMergeRects = rangeRects(view.layout, view.viewport, merge);
    expect(visibleMergeRects.length).toBeGreaterThan(1);
    const left = Math.min(...visibleMergeRects.map((rect) => rect.x));
    const top = Math.min(...visibleMergeRects.map((rect) => rect.y));
    const right = Math.max(...visibleMergeRects.map((rect) => rect.x + rect.w));
    const bottom = Math.max(...visibleMergeRects.map((rect) => rect.y + rect.h));
    const union = { x: left, y: top, w: right - left, h: bottom - top };
    expect(spy.clipRects).toContainEqual([union.x, union.y, union.w, union.h]);
    expect(spy.textDraws).toContainEqual(
      expect.objectContaining({ text: 'merged', x: union.x + union.w / 2 }),
    );
  });

  it('filters stored interior merge borders while preserving the perimeter', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } as const;
    mutators.mergeRange(store, merge);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        mutators.setCellFormat(
          store,
          { sheet: 0, row, col },
          {
            borders: { top: true, bottom: true, left: true, right: true },
          },
        );
      }
    }
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintBorders(renderer, state);

    // Four cells with all sides would draw 16 paths. Only the eight perimeter
    // sides remain after suppressing the four internal boundaries.
    expect(spy.strokes).toBe(8);
  });

  it('draws a uniform merged diagonal once across the merged surface', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } as const;
    const diagonal = { style: 'medium' as const, color: '#8f5b20' };
    mutators.mergeRange(store, merge);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        mutators.setCellFormat(
          store,
          { sheet: 0, row, col },
          { borders: { diagonalDown: diagonal, diagonalUp: diagonal } },
        );
      }
    }
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintBorders(renderer, state);

    // One down and one up line for the full merged surface. Drawing each of
    // four component cells would produce eight paths and visible crossings.
    expect(spy.strokes).toBe(2);
  });

  it('projects one merged diagonal through freeze and RTL quadrants', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 1, 1);
    mutators.setRightToLeft(store, true);
    store.setState((state) => ({
      ...state,
      viewport: {
        ...state.viewport,
        rowStart: 1,
        colStart: 1,
        rowCount: 4,
        colCount: 4,
        widthPx: 640,
      },
    }));
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 } as const;
    mutators.mergeRange(store, merge);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        mutators.setCellFormat(store, { sheet: 0, row, col }, { borders: { diagonalDown: true } });
      }
    }
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintBorders(renderer, state);

    const { view } = viewParts(state);
    const visibleMergeRects = rangeRects(view.layout, view.viewport, merge);
    expect(visibleMergeRects.length).toBeGreaterThan(1);
    // The global diagonal is clipped into the visible quadrants, so it is one
    // logical stroke rather than a mini-diagonal per freeze pane.
    expect(spy.strokes).toBe(1);
  });

  it('suppresses partial diagonal metadata inside a merge', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } as const;
    mutators.mergeRange(store, merge);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: merge.r0, col: merge.c0 },
      { borders: { diagonalDown: true } },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: merge.r1, col: merge.c1 },
      { borders: { diagonalUp: true } },
    );
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintBorders(renderer, state);

    // Excel drops non-uniform diagonal sides at merge time; stale loaded
    // metadata must not reintroduce those lines during rendering.
    expect(spy.strokes).toBe(0);
  });

  it('uses freeze and RTL-aware merge rectangles and ignores copy ranges on other sheets', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 1, 1);
    mutators.setRightToLeft(store, true);
    store.setState((state) => ({
      ...state,
      viewport: {
        ...state.viewport,
        rowStart: 1,
        colStart: 1,
        rowCount: 4,
        colCount: 4,
        widthPx: 640,
      },
    }));
    mutators.mergeRange(store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    mutators.setCopyRange(store, { sheet: 1, r0: 0, c0: 0, r1: 2, c1: 2 });
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintActive(renderer, state);

    const { view } = viewParts(state);
    const visibleMergeRects = rangeRects(view.layout, view.viewport, state.selection.range);
    expect(visibleMergeRects.length).toBeGreaterThan(1);
    // One active outline per visible freeze quadrant; no dashed [4,3] copy
    // marquee is emitted for a source belonging to another sheet.
    expect(spy.strokeRects).toHaveLength(visibleMergeRects.length);
    expect(spy.dashes).not.toContainEqual([4, 3]);
  });

  it('paints merged fill and pattern into every visible freeze quadrant', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 1, 1);
    mutators.setRightToLeft(store, true);
    store.setState((state) => ({
      ...state,
      viewport: {
        ...state.viewport,
        rowStart: 1,
        colStart: 1,
        rowCount: 4,
        colCount: 4,
        widthPx: 640,
      },
    }));
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 } as const;
    mutators.mergeRange(store, merge);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: merge.r0, col: merge.c0 },
      {
        fill: '#ffe0b0',
        fillPattern: 'horizontal',
        fillPatternColor: '#8f5b20',
      },
    );
    mutators.setActive(store, { sheet: 0, row: 2, col: 2 });
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintCells(renderer, state);

    const { view } = viewParts(state);
    const visibleMergeRects = rangeRects(view.layout, view.viewport, merge);
    expect(visibleMergeRects.length).toBeGreaterThan(1);
    // The custom fill must be emitted once for each visible quadrant. This
    // catches the old path that reused the first rect as content bounds for
    // all merged formatting.
    for (const rect of visibleMergeRects) {
      expect(spy.fills).toContainEqual({
        style: '#ffe0b0',
        rect: [rect.x, rect.y, rect.w, rect.h],
      });
    }
  });
});
