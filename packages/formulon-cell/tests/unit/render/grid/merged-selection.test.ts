import { describe, expect, it } from 'vitest';

import { addrKey } from '../../../../src/engine/address.js';
import type { Range } from '../../../../src/engine/types.js';
import {
  buildColLayout,
  buildRowLayout,
  layoutForView,
  rangeRects,
} from '../../../../src/render/geometry.js';
import type { GridHitState } from '../../../../src/render/grid/hit-state.js';
import { mergedDiagonalBorders } from '../../../../src/render/grid/merged-cells.js';
import { GridRenderer } from '../../../../src/render/grid.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type State,
} from '../../../../src/store/store.js';
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

class TrackingFormatMap extends Map<string, CellFormat> {
  reads = 0;

  constructor(
    source: ReadonlyMap<string, CellFormat>,
    private readonly trackedKeys: ReadonlySet<string>,
  ) {
    super(source);
  }

  override get(key: string): CellFormat | undefined {
    if (this.trackedKeys.has(key)) this.reads += 1;
    return super.get(key);
  }
}

function trackMergeFormatReads(state: State, merge: Range) {
  const trackedKeys = new Set<string>();
  for (let row = merge.r0; row <= merge.r1; row += 1) {
    for (let col = merge.c0; col <= merge.c1; col += 1) {
      trackedKeys.add(addrKey({ sheet: merge.sheet, row, col }));
    }
  }
  const formats = new TrackingFormatMap(state.format.formats, trackedKeys);
  return {
    state: { ...state, format: { ...state.format, formats } },
    formats,
  };
}

function hitsOf(renderer: GridRenderer): GridHitState {
  return (renderer as unknown as { hits: GridHitState }).hits;
}

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
    // The outline covers the merge's gridlines plus one pixel outside them.
    expect(spy.strokeRects[0]).toEqual([
      mergeRect.x,
      mergeRect.y,
      mergeRect.w + 1,
      mergeRect.h + 1,
    ]);
    // The selection outline is the same rectangle and must not be painted a
    // second time over the active merge outline. The chevron adds one stroke.
    expect(spy.strokeRects).toHaveLength(2);
    expect(hitsOf(renderer).validationChevron).toMatchObject({ row: 1, col: 1 });
    expect(hitsOf(renderer).validationChevron?.rect.x).toBe(mergeRect.x + mergeRect.w - 18);
    expect(hitsOf(renderer).fillHandle).toMatchObject({
      x: mergeRect.x + mergeRect.w - 3,
      y: mergeRect.y + mergeRect.h - 3,
    });
  });

  it('outlines a multi-cell range once and leaves the active cell without its own outline', () => {
    const store = createSpreadsheetStore();
    mutators.setActive(store, { sheet: 0, row: 1, col: 1 });
    mutators.extendRangeTo(store, { sheet: 0, row: 4, col: 3 });
    const state = store.getState();
    const spy = canvasSpy();
    const renderer = makeRenderer(state, spy);

    paintActive(renderer, state);

    const { view } = viewParts(state);
    const [rangeRect] = rangeRects(view.layout, view.viewport, state.selection.range);
    if (!rangeRect) throw new Error('expected the range to be visible');
    // A 1px sheet-coloured gap inside the border, then the border itself.
    expect(spy.strokeRects).toEqual([
      [rangeRect.x + 1.5, rangeRect.y + 1.5, rangeRect.w - 2, rangeRect.h - 2],
      [rangeRect.x, rangeRect.y, rangeRect.w + 1, rangeRect.h + 1],
    ]);
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

  it('skips diagonal metadata reads for merges outside the viewport', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 100, c0: 100, r1: 199, c1: 199 } as const;
    mutators.mergeRange(store, merge);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: merge.r0, col: merge.c0 },
      {
        borders: { diagonalDown: true },
      },
    );
    const tracked = trackMergeFormatReads(store.getState(), merge);
    const spy = canvasSpy();
    const renderer = makeRenderer(tracked.state, spy);

    paintBorders(renderer, tracked.state);

    // The merge covers 10,000 cells. Its diagonal summary must not read any
    // of them when the viewport cannot display the merged surface.
    expect(tracked.formats.reads).toBe(0);
  });

  it('reuses a merge diagonal summary while still painting visible perimeter cells', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } as const;
    mutators.mergeRange(store, merge);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        mutators.setCellFormat(
          store,
          { sheet: 0, row, col },
          {
            borders: { diagonalDown: true },
          },
        );
      }
    }
    const tracked = trackMergeFormatReads(store.getState(), merge);
    const spy = canvasSpy();
    const renderer = makeRenderer(tracked.state, spy);

    paintBorders(renderer, tracked.state);
    const readsAfterFirstPaint = tracked.formats.reads;
    paintBorders(renderer, tracked.state);

    // Each paint reads the four visible perimeter formats. Only the first
    // paint also needs four reads to establish diagonal uniformity.
    expect(readsAfterFirstPaint).toBe(8);
    expect(tracked.formats.reads - readsAfterFirstPaint).toBe(4);
    expect(spy.strokes).toBe(2);
  });

  it('recomputes a cached diagonal summary when the formats map changes', () => {
    const store = createSpreadsheetStore();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 } as const;
    mutators.mergeRange(store, merge);
    for (let row = merge.r0; row <= merge.r1; row += 1) {
      for (let col = merge.c0; col <= merge.c1; col += 1) {
        mutators.setCellFormat(
          store,
          { sheet: 0, row, col },
          {
            borders: { diagonalDown: true },
          },
        );
      }
    }
    const first = trackMergeFormatReads(store.getState(), merge);
    const spy = canvasSpy();
    const renderer = makeRenderer(first.state, spy);
    paintBorders(renderer, first.state);
    expect(spy.strokes).toBe(1);

    const changedFormats = new Map(first.state.format.formats);
    const firstCellKey = addrKey({ sheet: 0, row: merge.r0, col: merge.c0 });
    changedFormats.set(firstCellKey, { borders: {} });
    const changedState = {
      ...first.state,
      format: { ...first.state.format, formats: changedFormats },
    };
    const second = trackMergeFormatReads(changedState, merge);
    paintBorders(renderer, second.state);

    // A non-uniform diagonal is suppressed; reusing the previous Map's result
    // would incorrectly leave a stroke on the canvas.
    expect(spy.strokes).toBe(1);
    expect(second.formats.reads).toBe(8);
  });

  it('recomputes a cached diagonal summary when the merge object changes coordinates', () => {
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    const formats = new Map<string, CellFormat>([
      [addrKey({ sheet: 0, row: 0, col: 0 }), { borders: { diagonalDown: true } }],
    ]);

    expect(mergedDiagonalBorders(formats, merge)?.diagonalDown).toBe(true);
    merge.r1 = 1;

    expect(mergedDiagonalBorders(formats, merge)).toBeUndefined();
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

describe('GridRenderer while a chart or picture is selected', () => {
  it('hides the cell cursor, range outline, tint and fill handle', () => {
    const store = createSpreadsheetStore();
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 1, r1: 3, c1: 2 });
    const shown = canvasSpy();
    const shownRenderer = makeRenderer(store.getState(), shown);
    paintActive(shownRenderer, store.getState());
    paintCells(shownRenderer, store.getState());
    expect(shown.strokeRects.length).toBeGreaterThan(0);
    expect(shown.fills.some((fill) => fill.style === theme.accentSoft)).toBe(true);
    expect(hitsOf(shownRenderer).fillHandle).not.toBeNull();

    mutators.setSelectedObject(store, { id: 'chart-1', name: 'Chart 1' });
    const state = store.getState();
    const hidden = canvasSpy();
    const hiddenRenderer = makeRenderer(state, hidden);
    paintActive(hiddenRenderer, state);
    paintCells(hiddenRenderer, state);
    expect(hidden.strokeRects).toEqual([]);
    expect(hidden.fills.some((fill) => fill.style === theme.accentSoft)).toBe(false);
    // The renderer that painted a handle drops it once an object is selected.
    paintActive(shownRenderer, state);
    expect(hitsOf(shownRenderer).fillHandle).toBeNull();
  });
});
