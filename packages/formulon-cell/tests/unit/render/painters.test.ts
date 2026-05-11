import { describe, expect, it } from 'vitest';
import {
  FILL_HANDLE_SIZE,
  paintActiveCellOutline,
  paintCellBorders,
  paintCellText,
  paintCopyMarquee,
  paintFillHandle,
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
} {
  const fonts: string[] = [];
  const fills: Array<{ text: string; x: number; y: number; font: string }> = [];
  let font = '';
  const ctx = {
    get font(): string {
      return font;
    },
    set font(v: string) {
      font = v;
      fonts.push(v);
    },
    fillStyle: '',
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    rect(): void {},
    clip(): void {},
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
  return { ctx, fonts, fills };
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
    accent: '#0078d4',
    accentSoft: 'rgba(0,120,212,0.1)',
    cellErrorFg: '#c00',
    cellFormulaFg: '#333',
    cellBoolFg: '#060',
    cellNumFg: '#000',
    hoverStripe: 'rgba(0,0,0,0.04)',
    headerFg: '#666',
    headerFgActive: '#000',
    fontUi: 'sans-serif',
    fontMono: 'monospace',
    textCell: 13,
    textHeader: 11.5,
    ...over,
  }) as ResolvedTheme;

describe('paintFillHandle', () => {
  it('exports a 6px size constant', () => {
    expect(FILL_HANDLE_SIZE).toBe(6);
  });

  it('paints a white border ring then an accent square at the cell corner', () => {
    const { ctx, fills } = makeCtxSpy();
    // Cell rect at (10, 20), size 100×30. Bottom-right corner sits at (110, 50).
    const bounds = { x: 10, y: 20, w: 100, h: 30 };
    const hit = paintFillHandle(ctx, bounds, theme({ accent: '#0078d4' }));

    // Expect two fillRect calls — outer white halo, inner accent square.
    expect(fills).toHaveLength(2);
    const [halo, square] = fills;

    // Visible square is centred on the bottom-right corner; half-bleeds outside
    // the cell rect. With FILL_HANDLE_SIZE=6 the inner rect starts at
    //   x = 10 + 100 - 3 = 107
    //   y = 20 +  30 - 3 = 47
    expect(square?.style).toBe('#0078d4');
    expect(square?.rect).toEqual([107, 47, 6, 6]);

    // Halo wraps the square with a 1px white border.
    expect(halo?.style).toBe('#ffffff');
    expect(halo?.rect).toEqual([106, 46, 8, 8]);

    // Returned hit-rect spans the visible (halo) area for comfortable grabbing.
    expect(hit).toEqual({ x: 106, y: 46, w: 8, h: 8 });
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

describe('paintActiveCellOutline', () => {
  it('snaps the 2px active outline inside the cell bounds', () => {
    const { ctx, rects } = makeStrokeSpy();
    paintActiveCellOutline(
      ctx,
      { x: 10.2, y: 20.7, w: 80.4, h: 24.2 },
      theme({ accent: '#107c41' }),
    );

    expect(rects).toEqual([{ x: 11, y: 22, w: 78, h: 22, width: 2 }]);
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
    expect(spy.fills[0]?.x).toBe(7);
  });

  it('shrinks unformatted numbers to fit the cell width before falling back to ####', () => {
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

    // 686.666666666667 at 6px/char ≈ 96px. A 60px-wide cell (~46px available
    // after padding) can fit "686.667" (42px) but not the full string. The
    // renderer must trim, not clip.
    paintCellText({
      ctx,
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
    expect(rendered.length * 6).toBeLessThanOrEqual(60 - 14);
    expect(rendered).not.toMatch(/^#+$/);
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

    expect(fills).toEqual(['##']);
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

    expect(normal.fonts.at(-1)).toBe('400 13px "Times New Roman"');
    expect(bold.fonts.at(-1)).toBe('700 13px "Times New Roman"');
    expect(bold.fills[0]?.y).toBe(normal.fills[0]?.y);
  });

  it('uses stable text metrics for vertical placement', () => {
    expect(stableTextMetricsBox(13).ascent).toBeCloseTo(9.36);
    expect(stableTextMetricsBox(13).descent).toBeCloseTo(2.86);
  });
});
