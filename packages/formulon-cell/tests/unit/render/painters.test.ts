import { describe, expect, it } from 'vitest';
import {
  cellFont,
  cellFontFamily,
  shrinkFontSizeToFit,
} from '../../../src/render/painters/text.js';
import {
  FILL_HANDLE_SIZE,
  paintCellBorders,
  paintCellText,
  paintCopyMarquee,
  paintFillHandle,
  paintSelectionOutline,
  paintTableHeaderChevron,
  stableTextMetricsBox,
  textBaselineY,
} from '../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

/**
 * Minimal canvas spy. We only need to capture rect-fill calls + the active
 * `fillStyle` at each call so tests can assert paint order and colour.
 */
function makeCtxSpy(): {
  ctx: CanvasRenderingContext2D;
  fills: Array<{ style: string; rect: [number, number, number, number] }>;
} {
  const fills: Array<{ style: string; rect: [number, number, number, number] }> = [];
  let style = '';
  const ctx = {
    get fillStyle(): string {
      return style;
    },
    set fillStyle(v: string) {
      style = v;
    },
    fillRect(x: number, y: number, w: number, h: number): void {
      fills.push({ style, rect: [x, y, w, h] });
    },
    save(): void {},
    restore(): void {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fills };
}

function makeStrokeSpy(): {
  ctx: CanvasRenderingContext2D;
  strokes: Array<{ style: string; width: number; dash: number[] }>;
  rects: Array<{ x: number; y: number; w: number; h: number; width: number }>;
} {
  const strokes: Array<{ style: string; width: number; dash: number[] }> = [];
  const rects: Array<{ x: number; y: number; w: number; h: number; width: number }> = [];
  let strokeStyle = '';
  let lineWidth = 1;
  let dash: number[] = [];
  const ctx = {
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
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    moveTo(): void {},
    lineTo(): void {},
    setLineDash(v: number[]): void {
      dash = [...v];
    },
    stroke(): void {
      strokes.push({ style: strokeStyle, width: lineWidth, dash: [...dash] });
    },
    strokeRect(x: number, y: number, w: number, h: number): void {
      rects.push({ x, y, w, h, width: lineWidth });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, strokes, rects };
}

function makeTextSpy(): {
  ctx: CanvasRenderingContext2D;
  fonts: string[];
  fills: Array<{ text: string; x: number; y: number; font: string }>;
  directions: string[];
  strokeCount: () => number;
} {
  const fonts: string[] = [];
  const fills: Array<{ text: string; x: number; y: number; font: string }> = [];
  const directions: string[] = [];
  let strokes = 0;
  let font = '';
  let direction = 'inherit';
  const ctx = {
    get font(): string {
      return font;
    },
    set font(v: string) {
      font = v;
      fonts.push(v);
    },
    fillStyle: '',
    get direction(): string {
      return direction;
    },
    set direction(v: string) {
      direction = v;
      directions.push(v);
    },
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    rect(): void {},
    clip(): void {},
    moveTo(): void {},
    lineTo(): void {},
    stroke(): void {
      strokes += 1;
    },
    measureText(): TextMetrics {
      const bold = font.startsWith('700 ');
      return {
        width: 18,
        actualBoundingBoxAscent: bold ? 15 : 9,
        actualBoundingBoxDescent: bold ? 5 : 2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number): void {
      fills.push({ text, x, y, font });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, fonts, fills, directions, strokeCount: () => strokes };
}

const theme = (over: Partial<ResolvedTheme> = {}): ResolvedTheme =>
  ({
    bg: '#fff',
    bgRail: '#eee',
    bgElev: '#fff',
    bgHeader: '#eee',
    fg: '#000',
    fgMute: '#666',
    fgFaint: '#aaa',
    fgStrong: '#000',
    rule: '#ccc',
    ruleStrong: '#999',
    freezeRule: '#9b9b9b',
    accent: '#0078d4',
    accentSoft: 'rgba(0,120,212,0.1)',
    cellErrorFg: '#c00',
    cellFormulaFg: '#333',
    cellBoolFg: '#060',
    cellNumFg: '#000',
    hoverStripe: 'rgba(0,0,0,0.04)',
    headerBg: '#eee',
    headerRule: '#ccc',
    headerFg: '#666',
    headerFgActive: '#000',
    fontUi: 'sans-serif',
    fontMono: 'monospace',
    textCell: 13,
    textHeader: 11.5,
    ...over,
  }) as ResolvedTheme;

describe('paintFillHandle', () => {
  it('exports a 5px size constant', () => {
    expect(FILL_HANDLE_SIZE).toBe(5);
  });

  it('centres a 5px accent square on the trailing gridline corner with an inner white gap', () => {
    const { ctx, fills } = makeCtxSpy();
    // Bottom-right gridlines sit at x = 110, y = 50.
    const hit = paintFillHandle(ctx, { x: 10, y: 20, w: 100, h: 30 }, theme({ accent: '#0078d4' }));

    expect(fills).toHaveLength(2);
    const [gap, square] = fills;
    expect(square?.style).toBe('#0078d4');
    expect(square?.rect).toEqual([108, 48, 5, 5]);
    // The gap only faces into the selection (top and left).
    expect(gap?.style).toBe('#ffffff');
    expect(gap?.rect).toEqual([107, 47, 6, 6]);
    expect(hit).toEqual({ x: 107, y: 47, w: 7, h: 7 });
  });

  it('snaps fractional bounds and mirrors the gap in a right-to-left sheet', () => {
    const { ctx, fills } = makeCtxSpy();
    paintFillHandle(ctx, { x: 10.4, y: 20.4, w: 100, h: 29.8 }, theme({ accent: '#0078d4' }), true);
    expect(fills[1]?.rect).toEqual([8, 48, 5, 5]);
    expect(fills[0]?.rect).toEqual([8, 47, 6, 6]);
  });

  it('falls back to #0078d4 when theme.accent is empty', () => {
    const { ctx, fills } = makeCtxSpy();
    paintFillHandle(ctx, { x: 0, y: 0, w: 50, h: 20 }, theme({ accent: '' }));
    const square = fills[1];
    expect(square?.style).toBe('#0078d4');
  });
});

describe('paintTableHeaderChevron', () => {
  it('paints a compact dropdown affordance inside the cell header', () => {
    const fills: Array<{ style: string; rect?: [number, number, number, number] }> = [];
    let fillStyle = '';
    const ctx = {
      get fillStyle(): string {
        return fillStyle;
      },
      set fillStyle(v: string) {
        fillStyle = v;
      },
      strokeStyle: '',
      lineWidth: 1,
      save(): void {},
      restore(): void {},
      fillRect(x: number, y: number, w: number, h: number): void {
        fills.push({ style: fillStyle, rect: [x, y, w, h] });
      },
      strokeRect(): void {},
      beginPath(): void {},
      moveTo(): void {},
      lineTo(): void {},
      closePath(): void {},
      fill(): void {
        fills.push({ style: fillStyle });
      },
    } as unknown as CanvasRenderingContext2D;

    const hit = paintTableHeaderChevron(ctx, { x: 10, y: 20, w: 120, h: 24 }, theme());

    expect(hit).toEqual({ x: 113, y: 25, w: 14, h: 14 });
    expect(fills[0]).toEqual({ style: 'rgba(255,255,255,0.72)', rect: [113, 25, 14, 14] });
    expect(fills.at(-1)?.style).toBe(theme().fgMute);
  });
});

describe('paintCellBorders', () => {
  it('uses the spreadsheet-like automatic border color instead of the gridline color', () => {
    const { ctx, strokes } = makeStrokeSpy();
    paintCellBorders({
      ctx,
      bounds: { x: 0, y: 0, w: 80, h: 24 },
      theme: theme({ fgStrong: '#111111', ruleStrong: '#999999' }),
      value: { kind: 'blank' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { borders: { bottom: true } },
    });

    expect(strokes).toHaveLength(1);
    expect(strokes[0]?.style).toBe('#111111');
  });

  it('honors explicit border colors from cell format', () => {
    const { ctx, strokes } = makeStrokeSpy();
    paintCellBorders({
      ctx,
      bounds: { x: 0, y: 0, w: 80, h: 24 },
      theme: theme({ fgStrong: '#111111' }),
      value: { kind: 'blank' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { borders: { bottom: { style: 'thin', color: '#c00000' } } },
    });

    expect(strokes[0]?.style).toBe('#c00000');
  });
});

describe('paintSelectionOutline', () => {
  it('covers the range gridlines plus one pixel outside them', () => {
    const { ctx, rects } = makeStrokeSpy();
    // Gridlines snap to x = 10 / 91 and y = 21 / 45.
    paintSelectionOutline(
      ctx,
      { x: 10.2, y: 20.7, w: 80.4, h: 24.2 },
      theme({ accent: '#107c41' }),
    );

    expect(rects).toEqual([{ x: 10, y: 21, w: 82, h: 25, width: 2 }]);
  });
});

describe('paintCopyMarquee', () => {
  it('paints a black and white dashed copy marquee', () => {
    const { ctx, rects } = makeStrokeSpy();
    paintCopyMarquee(ctx, { x: 10, y: 20, w: 80, h: 24 });

    expect(rects).toEqual([
      { x: 11.5, y: 21.5, w: 77, h: 21, width: 1 },
      { x: 10.5, y: 20.5, w: 79, h: 23, width: 1 },
    ]);
  });
});

describe('textBaselineY', () => {
  it('keeps bottom-aligned text visually inside the cell padding', () => {
    const bounds = { x: 10, y: 20, w: 80, h: 20 };
    const box = { ascent: 10, descent: 3 };
    expect(textBaselineY(bounds, box, 'bottom', 4)).toBe(33);
  });

  it('centers the measured text box for middle alignment', () => {
    const bounds = { x: 0, y: 0, w: 80, h: 24 };
    const box = { ascent: 10, descent: 4 };
    expect(textBaselineY(bounds, box, 'middle', 4)).toBe(15);
  });
});

describe('paintCellText font strictness', () => {
  it('never enlarges a small font when shrink-to-fit is requested', () => {
    expect(shrinkFontSizeToFit(6.5, 100, 40)).toBeLessThanOrEqual(6.5);
  });
  it('keeps an explicit cell family first and appends the UI fallback stack', () => {
    expect(cellFontFamily('Calibri', 'Aptos, system-ui, sans-serif')).toBe(
      'Calibri, Aptos, system-ui, sans-serif',
    );
    expect(cellFontFamily('Calibri', 'var(--fc-font-ui, system-ui, sans-serif)')).toBe(
      'Calibri, var(--fc-font-ui, system-ui, sans-serif)',
    );
    expect(cellFontFamily(undefined, 'Aptos, system-ui, sans-serif')).toBe(
      'Aptos, system-ui, sans-serif',
    );
  });

  it('preserves the existing UI and formula monospace defaults without a cell family', () => {
    const uiTheme = theme({ fontUi: 'Aptos', fontMono: 'Menlo' });
    expect(cellFont(undefined, uiTheme, false).family).toBe('Aptos');
    expect(cellFont(undefined, uiTheme, true).family).toBe('Menlo');
  });

  it.each([
    [0.5, '400 6.5px sans-serif'],
    [1, '400 13px sans-serif'],
    [2, '400 26px sans-serif'],
    [4, '400 52px sans-serif'],
  ])('scales the painted glyph with viewport zoom %s', (zoom, expectedFont) => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 200 * zoom, h: 24 * zoom },
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: 'ABC' },
      formula: null,
      isActive: false,
      isInRange: false,
      zoom,
    });

    expect(spy.fonts.at(-1)).toBe(expectedFont);
  });

  it('applies the same reduced size and baseline direction to vertical alignment', () => {
    const normal = makeTextSpy();
    const superscript = makeTextSpy();
    const subscript = makeTextSpy();
    const base = {
      bounds: { x: 0, y: 0, w: 120, h: 30 },
      theme: theme({ textCell: 20 }),
      value: { kind: 'text' as const, value: 'x' },
      formula: null,
      isActive: false,
      isInRange: false,
    };
    paintCellText({ ...base, ctx: normal.ctx });
    paintCellText({ ...base, ctx: superscript.ctx, format: { fontVertAlign: 'superscript' } });
    paintCellText({ ...base, ctx: subscript.ctx, format: { fontVertAlign: 'subscript' } });

    expect(superscript.fonts.at(-1)).toBe('400 14px sans-serif');
    expect(subscript.fonts.at(-1)).toBe('400 14px sans-serif');
    expect(superscript.fills[0]?.y).toBeLessThan(normal.fills[0]?.y ?? Infinity);
    expect(subscript.fills[0]?.y).toBeGreaterThan(normal.fills[0]?.y ?? -Infinity);
  });

  it('uses the grid UI font for ordinary numbers like desktop spreadsheets', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 80, h: 20 },
      theme: theme({ fontUi: 'Aptos', fontMono: 'Menlo', textCell: 13 }),
      value: { kind: 'number', value: 1234 },
      formula: null,
      isActive: false,
      isInRange: false,
    });

    expect(spy.fonts.at(-1)).toBe('400 13px Aptos');
  });

  it('suppresses numeric zero text when the sheet hides zero values', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 80, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'number', value: 0 },
      formula: null,
      isActive: false,
      isInRange: false,
      showZeros: false,
    });

    expect(spy.fills).toEqual([]);
  });

  it('centers logical and error values by default', () => {
    const boolSpy = makeTextSpy();
    const errSpy = makeTextSpy();
    const base = {
      bounds: { x: 0, y: 0, w: 80, h: 20 },
      theme: theme({ textCell: 13 }),
      formula: null,
      isActive: false,
      isInRange: false,
    };

    paintCellText({ ...base, ctx: boolSpy.ctx, value: { kind: 'bool', value: true } });
    paintCellText({
      ...base,
      ctx: errSpy.ctx,
      value: { kind: 'error', code: 1, text: '#DIV/0!' },
    });

    expect(boolSpy.fills[0]?.x).toBe(40);
    expect(errSpy.fills[0]?.x).toBe(40);
  });

  it('applies explicit text direction to the canvas context', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 80, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: 'שלום' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { textDirection: 'rtl' },
    });

    expect(spy.directions).toContain('rtl');
  });

  it('uses hyperlink display text for blank hyperlinked cells', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 80, h: 20 },
      theme: theme({ textCell: 13, accent: '#0078d4' }),
      value: { kind: 'blank' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { hyperlink: 'https://example.test', hyperlinkDisplay: 'Example' },
    });

    expect(spy.fills[0]?.text).toBe('Example');
  });

  it.each([
    { underline: undefined, strike: false, strokes: 1 },
    { underline: false, strike: false, strokes: 0 },
    { underline: true, strike: false, strokes: 1 },
    { underline: 'double' as const, strike: false, strokes: 2 },
    { underline: false, strike: true, strokes: 1 },
  ])(
    'respects hyperlink underline $underline and strike $strike',
    ({ underline, strike, strokes }) => {
      const spy = makeTextSpy();
      paintCellText({
        ctx: spy.ctx,
        bounds: { x: 0, y: 0, w: 160, h: 24 },
        theme: theme({ textCell: 13 }),
        value: { kind: 'text', value: 'Example link' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { hyperlink: 'https://example.test', underline, strike },
      });
      expect(spy.fills[0]?.text).toBe('Example link');
      expect(spy.strokeCount()).toBe(strokes);
    },
  );

  it('applies Excel-style indent as left text inset', () => {
    const plain = makeTextSpy();
    const indented = makeTextSpy();
    const base = {
      bounds: { x: 10, y: 0, w: 100, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'text' as const, value: 'ABC' },
      formula: null,
      isActive: false,
      isInRange: false,
    };

    paintCellText({ ...base, ctx: plain.ctx });
    paintCellText({ ...base, ctx: indented.ctx, format: { indent: 2 } });

    expect(plain.fills[0]?.x).toBe(13);
    expect(indented.fills[0]?.x).toBe(29);
    expect((indented.fills[0]?.x ?? 0) - (plain.fills[0]?.x ?? 0)).toBe(16);
  });

  it('renders a phonetic guide above the base cell text', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 100, h: 24 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: '漢字' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { showPhonetic: true, phonetic: [{ start: 0, end: 2, text: 'かんじ' }] },
    });
    expect(spy.fills.map((fill) => fill.text)).toEqual(['カンジ', '漢字']);
    expect(spy.fills[0]?.font).toContain('7px');
    expect(spy.fills[0]?.y).toBeLessThan(spy.fills[1]?.y ?? Infinity);
  });

  it.each([undefined, false])(
    'draws no phonetic guide while Show Phonetic is %s',
    (showPhonetic) => {
      const spy = makeTextSpy();
      paintCellText({
        ctx: spy.ctx,
        bounds: { x: 0, y: 0, w: 100, h: 24 },
        theme: theme({ textCell: 13 }),
        value: { kind: 'text', value: '漢字' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { showPhonetic, phonetic: [{ start: 0, end: 2, text: 'かんじ' }] },
      });
      expect(spy.fills.map((fill) => fill.text)).toEqual(['漢字']);
    },
  );

  it('places each phonetic run over the span it annotates, in katakana by default', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 200, h: 24 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: '東京都' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: {
        showPhonetic: true,
        phonetic: [
          { start: 0, end: 2, text: 'とうきょう' },
          { start: 2, end: 3, text: 'と' },
        ],
      },
    });
    expect(spy.fills.map((fill) => fill.text)).toEqual(['トウキョウ', 'ト', '東京都']);
    // The second run annotates a later span, so its reading sits to the right.
    expect(spy.fills[0]?.x ?? 0).toBeLessThan(spy.fills[1]?.x ?? 0);
  });

  it('places a run with no usable span over the whole text', () => {
    const spy = makeTextSpy();
    const bounds = { x: 0, y: 0, w: 200, h: 24 };
    const base = {
      bounds,
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: '東京都' } as const,
      formula: null,
      isActive: false,
      isInRange: false,
    };
    const spanned = makeTextSpy();
    paintCellText({
      ...base,
      ctx: spy.ctx,
      format: { showPhonetic: true, phonetic: [{ start: 0, end: 0, text: 'とうきょうと' }] },
    });
    paintCellText({
      ...base,
      ctx: spanned.ctx,
      format: { showPhonetic: true, phonetic: [{ start: 0, end: 3, text: 'とうきょうと' }] },
    });
    expect(spy.fills[0]?.x).toBe(spanned.fills[0]?.x);
  });

  it('uses the monospace font only for formula-display mode', () => {
    const spy = makeTextSpy();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 80, h: 20 },
      theme: theme({ fontUi: 'Aptos', fontMono: 'Menlo', textCell: 13 }),
      value: { kind: 'number', value: 2 },
      formula: '=A1+1',
      showFormulas: true,
      isActive: false,
      isInRange: false,
    });

    expect(spy.fonts.at(-1)).toBe('400 13px Menlo');
    expect(spy.fills[0]?.text).toBe('=A1+1');
    expect(spy.fills[0]?.x).toBe(3);
  });

  it.each([false, true])('fits General numbers before #### (explicit=%s)', (explicit) => {
    const fills: string[] = [];
    const ctx = {
      font: '',
      fillStyle: '',
      textBaseline: 'alphabetic',
      textAlign: 'left',
      save(): void {},
      restore(): void {},
      beginPath(): void {},
      rect(): void {},
      clip(): void {},
      measureText(text: string): TextMetrics {
        return { width: text.length * 6 } as TextMetrics;
      },
      fillText(text: string): void {
        fills.push(text);
      },
    } as unknown as CanvasRenderingContext2D;

    // 686.666666666667 at 6px/char ≈ 96px. A 60px-wide cell (~54px available
    // after padding) can fit "686.6667" (48px) but not the full string. The
    // renderer must trim, not clip.
    paintCellText({
      ctx,
      format: explicit ? { numFmt: { kind: 'general' } } : undefined,
      bounds: { x: 0, y: 0, w: 60, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'number', value: 686.6666666666667 },
      formula: null,
      isActive: false,
      isInRange: false,
    });

    expect(fills).toHaveLength(1);
    const rendered = fills[0] ?? '';
    expect(rendered.startsWith('686')).toBe(true);
    expect(rendered.length * 6).toBeLessThanOrEqual(60 - 6);
    expect(rendered).not.toMatch(/^#+$/);
  });

  it('uses the locale decimal mark in width-fitting scientific output', () => {
    const fills: string[] = [];
    const ctx = {
      font: '',
      fillStyle: '',
      direction: 'inherit',
      textBaseline: 'alphabetic',
      textAlign: 'left',
      save(): void {},
      restore(): void {},
      beginPath(): void {},
      rect(): void {},
      clip(): void {},
      measureText(text: string): TextMetrics {
        return { width: text.length * 6 } as TextMetrics;
      },
      fillText(text: string): void {
        fills.push(text);
      },
    } as unknown as CanvasRenderingContext2D;

    paintCellText({
      ctx,
      bounds: { x: 0, y: 0, w: 66, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'number', value: 123456789012 },
      formula: null,
      isActive: false,
      isInRange: false,
      locale: 'de-DE',
    });

    expect(fills).toEqual(['1,2346E+11']);
  });

  it('falls back to #### when the integer part itself is wider than the cell', () => {
    const fills: string[] = [];
    const ctx = {
      font: '',
      fillStyle: '',
      textBaseline: 'alphabetic',
      textAlign: 'left',
      save(): void {},
      restore(): void {},
      beginPath(): void {},
      rect(): void {},
      clip(): void {},
      measureText(text: string): TextMetrics {
        return { width: text.length * 6 } as TextMetrics;
      },
      fillText(text: string): void {
        fills.push(text);
      },
    } as unknown as CanvasRenderingContext2D;

    paintCellText({
      ctx,
      bounds: { x: 0, y: 0, w: 24, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'number', value: 123456789.0123 },
      formula: null,
      isActive: false,
      isInRange: false,
    });

    expect(fills).toHaveLength(1);
    expect(fills[0]).toMatch(/^#+$/);
  });

  it('renders hashes for formatted numbers that do not fit the cell width', () => {
    const fills: string[] = [];
    const ctx = {
      font: '',
      fillStyle: '',
      textBaseline: 'alphabetic',
      textAlign: 'left',
      save(): void {},
      restore(): void {},
      beginPath(): void {},
      rect(): void {},
      clip(): void {},
      measureText(text: string): TextMetrics {
        return { width: text.length * 6 } as TextMetrics;
      },
      fillText(text: string): void {
        fills.push(text);
      },
    } as unknown as CanvasRenderingContext2D;

    paintCellText({
      ctx,
      bounds: { x: 0, y: 0, w: 28, h: 20 },
      theme: theme({ textCell: 13 }),
      value: { kind: 'number', value: 123456 },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { numFmt: { kind: 'fixed', decimals: 0, thousands: true } },
    });

    expect(fills).toEqual(['###']);
  });

  it('keeps the same font size and baseline when only bold changes', () => {
    const bounds = { x: 0, y: 0, w: 80, h: 20 };
    const normal = makeTextSpy();
    const bold = makeTextSpy();

    paintCellText({
      ctx: normal.ctx,
      bounds,
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: 'ABC' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { fontFamily: 'Times New Roman', fontSize: 13 },
    });
    paintCellText({
      ctx: bold.ctx,
      bounds,
      theme: theme({ textCell: 13 }),
      value: { kind: 'text', value: 'ABC' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { bold: true, fontFamily: 'Times New Roman', fontSize: 13 },
    });

    expect(normal.fonts.at(-1)).toBe('400 13px "Times New Roman", sans-serif');
    expect(bold.fonts.at(-1)).toBe('700 13px "Times New Roman", sans-serif');
    expect(bold.fills[0]?.y).toBe(normal.fills[0]?.y);
  });

  it('draws two lines for double and accounting-double underlines', () => {
    for (const underline of ['double', 'doubleAccounting'] as const) {
      const spy = makeTextSpy();
      paintCellText({
        ctx: spy.ctx,
        bounds: { x: 0, y: 0, w: 80, h: 20 },
        theme: theme({ textCell: 13 }),
        value: { kind: 'text', value: 'ABC' },
        formula: null,
        isActive: false,
        isInRange: false,
        format: { underline },
      });
      expect(spy.strokeCount()).toBe(2);
    }
  });

  it('uses stable text metrics for vertical placement', () => {
    expect(stableTextMetricsBox(13).ascent).toBeCloseTo(9.36);
    expect(stableTextMetricsBox(13).descent).toBeCloseTo(2.86);
  });
});

describe('paintCellText font effects', () => {
  const effectSpy = () => {
    const strokes: string[] = [];
    const fills: string[] = [];
    const shadows: Array<{ color: string; x: number; y: number }> = [];
    let shadowColor = '';
    let shadowOffsetX = 0;
    let shadowOffsetY = 0;
    const ctx = makeTextSpy().ctx;
    Object.defineProperties(ctx, {
      shadowColor: {
        get: () => shadowColor,
        set: (v: string) => {
          shadowColor = v;
        },
      },
      shadowOffsetX: {
        get: () => shadowOffsetX,
        set: (v: number) => {
          shadowOffsetX = v;
        },
      },
      shadowOffsetY: {
        get: () => shadowOffsetY,
        set: (v: number) => {
          shadowOffsetY = v;
        },
      },
    });
    Object.assign(ctx, {
      shadowBlur: 0,
      lineWidth: 1,
      strokeStyle: '',
      fillText(text: string): void {
        fills.push(text);
        shadows.push({ color: shadowColor, x: shadowOffsetX, y: shadowOffsetY });
      },
      strokeText(text: string): void {
        strokes.push(text);
      },
    });
    return { ctx, strokes, fills, shadows };
  };
  const base = {
    bounds: { x: 0, y: 0, w: 120, h: 30 },
    theme: theme(),
    value: { kind: 'text' as const, value: 'Hi' },
    formula: null,
    isActive: false,
    isInRange: false,
  };

  it('paints outlined text as a hollow stroke', () => {
    const spy = effectSpy();
    paintCellText({ ...base, ctx: spy.ctx, fontEffects: true, format: { fontOutline: true } });
    expect(spy.strokes).toEqual(['Hi']);
    expect(spy.fills).toEqual([]);
  });

  it('paints shadowed text with an offset shadow', () => {
    const spy = effectSpy();
    paintCellText({ ...base, ctx: spy.ctx, fontEffects: true, format: { fontShadow: true } });
    expect(spy.fills).toEqual(['Hi']);
    expect(spy.shadows[0]?.color).not.toBe('');
    expect(spy.shadows[0]?.x).toBeGreaterThan(0);
    expect(spy.shadows[0]?.y).toBeGreaterThan(0);
  });

  it('paints the effects as plain text outside the Mac profile', () => {
    const spy = effectSpy();
    paintCellText({ ...base, ctx: spy.ctx, format: { fontOutline: true, fontShadow: true } });
    expect(spy.strokes).toEqual([]);
    expect(spy.fills).toEqual(['Hi']);
    expect(spy.shadows[0]?.x ?? 0).toBe(0);
  });

  it('paints plain text without a shadow', () => {
    const spy = effectSpy();
    paintCellText({ ...base, ctx: spy.ctx });
    expect(spy.shadows[0]?.x ?? 0).toBe(0);
  });
});
