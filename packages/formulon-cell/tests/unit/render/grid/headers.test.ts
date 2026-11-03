import { describe, expect, it } from 'vitest';
import { buildColLayout, buildRowLayout, layoutForView } from '../../../../src/render/geometry.js';
import { paintHeaders } from '../../../../src/render/grid/headers.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const theme = (over: Partial<ResolvedTheme> = {}): ResolvedTheme =>
  ({
    bg: '#fff',
    bgRail: '#faf9f8',
    bgElev: '#fff',
    bgHeader: '#e6e6e6',
    fg: '#201f1e',
    fgMute: '#605e5c',
    fgFaint: '#8a8886',
    fgStrong: '#11100f',
    rule: '#d9d9d9',
    ruleStrong: '#bfbfbf',
    freezeRule: '#9b9b9b',
    accent: '#107c41',
    accentFg: '#fff',
    accentSoft: 'rgba(16,124,65,0.12)',
    cellErrorFg: '#a4262c',
    cellFormulaFg: '#201f1e',
    cellBoolFg: '#107c41',
    cellNumFg: '#11100f',
    hoverStripe: 'rgba(16,124,65,0.025)',
    headerBg: '#faf9f8',
    headerRule: '#d9d9d9',
    headerFg: '#605e5c',
    headerFgActive: '#107c41',
    fontUi: 'Aptos, sans-serif',
    fontMono: 'ui-monospace, monospace',
    textCell: 13,
    textHeader: 11.5,
    ...over,
  }) as ResolvedTheme;

function makeCtxSpy(): {
  ctx: CanvasRenderingContext2D;
  fills: Array<{ style: string; alpha: number; rect?: [number, number, number, number] }>;
  texts: Array<{
    text: string;
    font: string;
    style: string;
    x: number;
    y: number;
    align: string;
    baseline: string;
  }>;
  strokes: string[];
} {
  const fills: Array<{ style: string; alpha: number; rect?: [number, number, number, number] }> =
    [];
  const texts: Array<{
    text: string;
    font: string;
    style: string;
    x: number;
    y: number;
    align: string;
    baseline: string;
  }> = [];
  const strokes: string[] = [];
  let fillStyle = '';
  let strokeStyle = '';
  let lineWidth = 1;
  let font = '';
  let alpha = 1;
  const ctx = {
    get fillStyle(): string {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    get strokeStyle(): string {
      return strokeStyle;
    },
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    get lineWidth(): number {
      return lineWidth;
    },
    set lineWidth(v: number) {
      lineWidth = v;
    },
    get font(): string {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    get globalAlpha(): number {
      return alpha;
    },
    set globalAlpha(v: number) {
      alpha = v;
    },
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {
      alpha = 1;
    },
    beginPath(): void {},
    closePath(): void {},
    moveTo(): void {},
    lineTo(): void {},
    quadraticCurveTo(): void {},
    stroke(): void {
      strokes.push(strokeStyle);
    },
    fill(): void {
      fills.push({ style: fillStyle, alpha });
    },
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style: fillStyle, alpha, rect: [x, y, w, h] });
    },
    fillText(
      this: { textAlign: string; textBaseline: string },
      text: string,
      x: number,
      y: number,
    ): void {
      texts.push({
        text,
        font,
        style: fillStyle,
        x,
        y,
        align: this.textAlign,
        baseline: this.textBaseline,
      });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills, texts, strokes };
}

describe('paintHeaders', () => {
  it('centres row numbers in the strip and sets labels in the given family', () => {
    const state = createSpreadsheetStore().getState();
    state.viewport = { rowStart: 0, rowCount: 12, colStart: 0, colCount: 2, zoom: 1, widthPx: 320 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, texts } = makeCtxSpy();
    const view = { ...state, layout: layoutForView(state) };

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 320, cssHeight: 400 },
      view,
      theme(),
      cols,
      rows,
      'Courier New, sans-serif',
    );

    // The strip ends in a 1px rule at the data edge; labels centre on the rest.
    const centre = view.layout.headerColWidth / 2 - 0.5;
    for (const label of ['1', '9', '10', '12']) {
      expect(texts.find((t) => t.text === label)).toMatchObject({ x: centre, align: 'center' });
    }
    expect(texts.find((t) => t.text === 'A')?.font).toContain('Courier New, sans-serif');
    expect(texts.find((t) => t.text === '10')?.font).toContain('Courier New, sans-serif');
  });

  it('draws the accent bar on every selected column and row header', () => {
    const state = createSpreadsheetStore().getState();
    state.selection = {
      active: { sheet: 0, row: 1, col: 1 },
      range: { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 },
      anchor: { sheet: 0, row: 1, col: 1 },
      extraRanges: [],
    };
    state.viewport = { rowStart: 0, rowCount: 4, colStart: 0, colCount: 4, zoom: 1, widthPx: 320 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, fills } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 320, cssHeight: 160 },
      { ...state, layout: layoutForView(state) },
      theme(),
      cols,
      rows,
    );

    const bars = fills.filter((f) => f.style === '#107c41' && f.rect);
    const colBars = bars.filter((f) => f.rect?.[3] === 2 && (f.rect?.[2] ?? 0) > 2);
    const rowBars = bars.filter((f) => f.rect?.[2] === 2 && (f.rect?.[3] ?? 0) > 2);
    expect(colBars).toHaveLength(2);
    expect(rowBars).toHaveLength(2);
  });

  it('uses regular idle header text and semibold selected header text', () => {
    const state = createSpreadsheetStore().getState();
    state.selection = {
      active: { sheet: 0, row: 1, col: 1 },
      range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      anchor: { sheet: 0, row: 1, col: 1 },
      extraRanges: [],
    };
    state.viewport = { rowStart: 0, rowCount: 3, colStart: 0, colCount: 3, zoom: 1, widthPx: 260 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, texts } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 260, cssHeight: 140 },
      { ...state, layout: layoutForView(state) },
      theme(),
      cols,
      rows,
    );

    expect(texts.find((entry) => entry.text === 'A')).toMatchObject({
      font: '400 11.5px Aptos, sans-serif',
      style: '#605e5c',
    });
    expect(texts.find((entry) => entry.text === 'B')).toMatchObject({
      font: '600 11.5px Aptos, sans-serif',
      style: '#107c41',
    });
    expect(texts.find((entry) => entry.text === '2')).toMatchObject({
      font: '600 11.5px Aptos, sans-serif',
      style: '#107c41',
    });
  });

  it('drops the header highlight while a chart or picture is selected', () => {
    const state = createSpreadsheetStore().getState();
    state.selection = {
      active: { sheet: 0, row: 1, col: 1 },
      range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      anchor: { sheet: 0, row: 1, col: 1 },
      extraRanges: [],
    };
    state.ui = { ...state.ui, selectedObject: { id: 'chart-1', name: 'Chart 1' } };
    state.viewport = { rowStart: 0, rowCount: 3, colStart: 0, colCount: 3, zoom: 1, widthPx: 260 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, texts, fills, strokes } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 260, cssHeight: 140 },
      { ...state, layout: layoutForView(state) },
      theme(),
      cols,
      rows,
    );

    for (const label of ['B', '2']) {
      expect(texts.find((entry) => entry.text === label)).toMatchObject({
        font: '400 11.5px Aptos, sans-serif',
        style: '#605e5c',
      });
    }
    expect(fills.some((fill) => fill.style === '#e6e6e6')).toBe(false);
    expect(fills.some((fill) => fill.style === '#107c41')).toBe(false);
    expect(strokes).not.toContain('#107c41');

    // The same cell without a selected object carries the accent markers.
    const shown = makeCtxSpy();
    paintHeaders(
      { ctx: shown.ctx, dpr: 1, cssWidth: 260, cssHeight: 140 },
      { ...state, ui: { ...state.ui, selectedObject: null }, layout: layoutForView(state) },
      theme(),
      cols,
      rows,
    );
    // The accent bars sit on the two pixels just before the header edge.
    const view = layoutForView(state);
    const bars = shown.fills.filter((fill) => fill.style === '#107c41').map((fill) => fill.rect);
    const oy = Math.round(view.headerRowHeight);
    const ox = Math.round(view.headerColWidth);
    expect(bars.some((rect) => rect?.[1] === oy - 2 && rect[3] === 2)).toBe(true);
    expect(bars.some((rect) => rect?.[0] === ox - 2 && rect[2] === 2)).toBe(true);
  });

  it('parks the row rail and the corner chip on the right for a right-to-left sheet', () => {
    const state = createSpreadsheetStore().getState();
    state.ui = { ...state.ui, rightToLeft: true };
    state.viewport = { rowStart: 0, rowCount: 3, colStart: 0, colCount: 3, zoom: 1, widthPx: 260 };
    const layout = layoutForView(state);
    const cols = buildColLayout(layout, state.viewport);
    const rows = buildRowLayout(layout, state.viewport);
    const { ctx, fills } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 260, cssHeight: 140 },
      { ...state, layout },
      theme(),
      cols,
      rows,
    );

    const ox = layout.outlineRowGutter + layout.headerColWidth;
    // Corner chip: full rail width, flush with the right edge.
    expect(fills[0]?.rect).toEqual([
      260 - ox,
      0,
      ox,
      layout.outlineColGutter + layout.headerRowHeight,
    ]);
    // Row rail: same left edge, running the full height below the column rail.
    expect(fills.some((f) => f.rect?.[0] === 260 - ox && f.rect?.[2] === ox)).toBe(true);
  });

  it('keeps the select-all corner marker subtle', () => {
    const state = createSpreadsheetStore().getState();
    state.viewport = { rowStart: 0, rowCount: 1, colStart: 0, colCount: 1, zoom: 1, widthPx: 260 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, fills } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 1, cssWidth: 120, cssHeight: 80 },
      { ...state, layout: layoutForView(state) },
      theme(),
      cols,
      rows,
    );

    expect(fills[1]).toMatchObject({ style: '#605e5c', alpha: 0.34 });
  });

  it('draws a sort arrow instead of the chevron on a sorted filter column', () => {
    const paint = (sorted: boolean) => {
      const state = createSpreadsheetStore().getState();
      const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      state.ui = {
        ...state.ui,
        filterRanges: [range],
        filterSorts: sorted ? [{ range, keys: [{ byCol: 1, direction: 'desc' }] }] : [],
      };
      state.viewport = {
        rowStart: 0,
        rowCount: 4,
        colStart: 0,
        colCount: 3,
        zoom: 1,
        widthPx: 320,
      };
      const cols = buildColLayout(state.layout, state.viewport);
      const rows = buildRowLayout(state.layout, state.viewport);
      const spy = makeCtxSpy();
      paintHeaders(
        { ctx: spy.ctx, dpr: 1, cssWidth: 320, cssHeight: 160 },
        { ...state, layout: layoutForView(state) },
        theme(),
        cols,
        rows,
      );
      return spy;
    };
    const plain = paint(false);
    const sorted = paint(true);
    // One chevron fill disappears and one arrow stroke appears for column B.
    expect(sorted.fills.length).toBe(plain.fills.length - 1);
    expect(sorted.strokes.length).toBe(plain.strokes.length + 1);
  });

  it('sets row numbers on a baseline just above the row bottom, as Excel does on tall rows', () => {
    // Excel (DPR 2): a 48 px row puts the digit ink at y 26..41 and a 100 px
    // row at 174..189, both 6 device px above the row's bottom edge.
    const state = createSpreadsheetStore().getState();
    state.layout.rowHeights = new Map([
      [0, 24],
      [2, 50],
    ]);
    state.viewport = { rowStart: 0, rowCount: 4, colStart: 0, colCount: 2, zoom: 1, widthPx: 320 };
    const layout = layoutForView(state);
    const cols = buildColLayout(layout, state.viewport);
    const rows = buildRowLayout(layout, state.viewport);
    const { ctx, texts } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 2, cssWidth: 320, cssHeight: 400 },
      { ...state, layout },
      theme(),
      cols,
      rows,
    );

    const oy = layout.headerRowHeight;
    expect(texts.find((t) => t.text === '1')).toMatchObject({
      y: oy + 24 - 3,
      baseline: 'alphabetic',
    });
    expect(texts.find((t) => t.text === '3')).toMatchObject({
      y: oy + 24 + 20 + 50 - 3,
      baseline: 'alphabetic',
    });
  });

  it('centres column letters on the header cell less its trailing separator', () => {
    // Excel (DPR 2): column A spans x 0..119 with its separator on 118..119;
    // the letter's ink centres on 58.5, the middle of 0..117.
    const state = createSpreadsheetStore().getState();
    state.layout.colWidths = new Map([[0, 60]]);
    state.viewport = { rowStart: 0, rowCount: 2, colStart: 0, colCount: 2, zoom: 1, widthPx: 320 };
    const layout = layoutForView(state);
    const cols = buildColLayout(layout, state.viewport);
    const rows = buildRowLayout(layout, state.viewport);
    const { ctx, texts } = makeCtxSpy();

    paintHeaders(
      { ctx, dpr: 2, cssWidth: 320, cssHeight: 200 },
      { ...state, layout },
      theme(),
      cols,
      rows,
    );

    const ox = layout.headerColWidth;
    expect(texts.find((t) => t.text === 'A')).toMatchObject({ x: ox + 29.5, align: 'center' });
    expect(texts.find((t) => t.text === 'B')).toMatchObject({ x: ox + 60 + 31.5, align: 'center' });
  });
});
