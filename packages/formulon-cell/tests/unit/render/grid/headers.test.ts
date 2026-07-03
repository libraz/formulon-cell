import { describe, expect, it } from 'vitest';
import { buildColLayout, buildRowLayout } from '../../../../src/render/geometry.js';
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
} {
  const fills: Array<{ style: string; alpha: number; rect?: [number, number, number, number] }> =
    [];
  const texts: Array<{ text: string; font: string; style: string }> = [];
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
    stroke(): void {},
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
  return { ctx, fills, texts };
}

describe('paintHeaders', () => {
  it('uses regular idle header text and semibold selected header text', () => {
    const state = createSpreadsheetStore().getState();
    state.selection = {
      active: { sheet: 0, row: 1, col: 1 },
      range: { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
      anchor: { sheet: 0, row: 1, col: 1 },
      extraRanges: [],
    };
    state.viewport = { rowStart: 0, rowCount: 3, colStart: 0, colCount: 3, zoom: 1 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, texts } = makeCtxSpy();

    paintHeaders({ ctx, dpr: 1, cssWidth: 260, cssHeight: 140 }, state, theme(), cols, rows);

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

  it('keeps the select-all corner marker subtle', () => {
    const state = createSpreadsheetStore().getState();
    state.viewport = { rowStart: 0, rowCount: 1, colStart: 0, colCount: 1, zoom: 1 };
    const cols = buildColLayout(state.layout, state.viewport);
    const rows = buildRowLayout(state.layout, state.viewport);
    const { ctx, fills } = makeCtxSpy();

    paintHeaders({ ctx, dpr: 1, cssWidth: 120, cssHeight: 80 }, state, theme(), cols, rows);

    expect(fills[1]).toMatchObject({ style: '#605e5c', alpha: 0.34 });
  });
});
