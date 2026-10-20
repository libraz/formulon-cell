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
    accent: '#107c41',
    accentFg: '#fff',
    accentSoft: 'rgba(16,124,65,0.12)',
    cellErrorFg: '#a4262c',
    cellFormulaFg: '#201f1e',
    cellBoolFg: '#107c41',
    cellNumFg: '#11100f',
    hoverStripe: 'rgba(16,124,65,0.025)',
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
  texts: Array<{ text: string; font: string; style: string }>;
  strokes: string[];
} {
  const fills: Array<{ style: string; alpha: number; rect?: [number, number, number, number] }> =
    [];
  const texts: Array<{ text: string; font: string; style: string }> = [];
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
    stroke(): void {
      strokes.push(strokeStyle);
    },
    fill(): void {
      fills.push({ style: fillStyle, alpha });
    },
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style: fillStyle, alpha, rect: [x, y, w, h] });
    },
    fillText(text: string): void {
      texts.push({ text, font, style: fillStyle });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills, texts, strokes };
}

describe('paintHeaders', () => {
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
    expect(shown.strokes).toContain('#107c41');
    expect(shown.fills.some((fill) => fill.style === '#107c41')).toBe(true);
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
});
