import { describe, expect, it } from 'vitest';
import { paintCellText } from '../../../src/render/painters/text.js';
import {
  cellTextWidth,
  drawCellTextRun,
  paintedFontSize,
  wholePointAdvances,
} from '../../../src/render/text-metrics.js';
import type { CellFormat } from '../../../src/store/store.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

/** Arial advances in font units (unitsPerEm 2048), from the face's hmtx. */
const ARIAL_UNITS: Record<string, number> = {
  ' ': 569,
  '0': 1139,
  A: 1366,
  H: 1479,
  R: 1479,
  V: 1366,
  W: 1933,
  a: 1139,
  d: 1139,
  e: 1139,
  f: 569,
  g: 1139,
  h: 1139,
  i: 455,
  l: 455,
  n: 1139,
  o: 1139,
  r: 682,
  t: 569,
  v: 1024,
  w: 1479,
  x: 1024,
};

/** Arial's 'AV' pair kern, which the canvas applies to a whole run. */
const AV_KERN_UNITS = -152;

/** Canvas spy measuring Arial: fractional advances and pair kerning for a
 *  run, as a browser canvas reports them. */
function arialCanvas() {
  const fills: Array<{ text: string; x: number; y: number }> = [];
  const translates: [number, number][] = [];
  let font = '400 11px Arial';
  const size = (): number => Number(font.match(/(\d+(?:\.\d+)?)px/)?.[1] ?? 11);
  const ctx = {
    get font(): string {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    direction: 'inherit',
    textBaseline: 'alphabetic',
    textAlign: 'left',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    rect(): void {},
    clip(): void {},
    translate(x: number, y: number): void {
      translates.push([x, y]);
    },
    rotate(): void {},
    moveTo(): void {},
    lineTo(): void {},
    stroke(): void {},
    measureText(text: string): TextMetrics {
      const chars = [...text];
      let units = 0;
      chars.forEach((ch, i) => {
        units += ARIAL_UNITS[ch] ?? 1139;
        if (ch === 'A' && chars[i + 1] === 'V') units += AV_KERN_UNITS;
        if (ch === 'V' && chars[i + 1] === 'A') units += AV_KERN_UNITS;
      });
      const s = size();
      return {
        width: (units / 2048) * s,
        fontBoundingBoxAscent: (s * 10) / 11,
        fontBoundingBoxDescent: (s * 2) / 11,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number): void {
      fills.push({ text, x, y });
    },
  };
  return {
    ctx: ctx as unknown as CanvasRenderingContext2D,
    fills,
    translates,
    setFont: (f: string) => (font = f),
  };
}

const widthAt = (text: string, px: number, wholePoint = true): number => {
  const { ctx } = arialCanvas();
  ctx.font = `400 ${px}px Arial`;
  return cellTextWidth(ctx, text, wholePoint);
};

describe('whole-point advances', () => {
  it('applies on the Mac point model only', () => {
    expect(wholePointAdvances({ screenScale: 0.75 })).toBe(true);
    expect(wholePointAdvances({ screenScale: 1 })).toBe(false);
    expect(wholePointAdvances({})).toBe(false);
  });

  // Predicted widths from the measured rule; Excel's ink widths (device px,
  // DPR 2) sit 0-3 px inside them: 113, 39, 61, 201, 211, 109, 18.
  it('sums advances rounded at the effective size, with no kerning, at 11 pt', () => {
    expect(widthAt('AVAVAVAV', 11)).toBe(56);
    expect(widthAt('iiiiiiiiii', 11)).toBe(20);
    expect(widthAt('ffffffffff', 11)).toBe(30);
    expect(widthAt('WWWWWWWWWW', 11)).toBe(100);
    expect(widthAt('Right aligned overflow', 11)).toBe(106);
    expect(widthAt('aligned text', 11)).toBe(55);
    expect(widthAt('fill', 11)).toBe(9);
  });

  it('rounds at 8 and 20 pt', () => {
    expect(widthAt('iiiiiiiiii', 8)).toBe(20);
    expect(widthAt('WWWWWWWWWW', 8)).toBe(80);
    expect(widthAt('aligned text', 8)).toBe(38);
    expect(widthAt('ffffffffff', 20)).toBe(60);
    expect(widthAt('AVAVAVAV', 20)).toBe(104);
    expect(widthAt('WWWWWWWWWW', 20)).toBe(190);
  });

  it('rounds at the zoomed size, not at 100% then scaled', () => {
    // Arial 11 at 150% paints at 17 px: Excel ink 76, 101, 320, 176 device px.
    expect(paintedFontSize((11 * 4) / 3, 1.5 * 0.75, true)).toBe(17);
    expect(widthAt('iiiiiiiiii', 17)).toBe(40);
    expect(widthAt('ffffffffff', 17)).toBe(50);
    expect(widthAt('WWWWWWWWWW', 17)).toBe(160);
    expect(widthAt('AVAVAVAV', 17)).toBe(88);
  });

  it('snaps the zoomed size to a whole px', () => {
    // 75%: 'Hxg 123' Excel ink 55 device px = 28 px of advances at 8 px.
    expect(paintedFontSize((11 * 4) / 3, 0.75 * 0.75, true)).toBe(8);
    expect(paintedFontSize((11 * 4) / 3, 1.25 * 0.75, true)).toBe(14);
    expect(paintedFontSize((11 * 4) / 3, 0.75 * 0.75, false)).toBeCloseTo(8.25);
  });

  it('keeps the canvas run width off the point model', () => {
    expect(widthAt('AVAVAVAV', 11, false)).toBeCloseTo(
      ((4 * 1366 + 4 * 1366 - 7 * 152) / 2048) * 11,
    );
  });

  it('keeps shaped runs whole', () => {
    expect(widthAt('e\u0301', 11)).toBeCloseTo(((1139 + 1139) / 2048) * 11);
  });

  it('places each glyph at the sum of the rounded advances from a whole px', () => {
    const { ctx, fills } = arialCanvas();
    ctx.font = '400 11px Arial';
    ctx.textAlign = 'right';
    const left = drawCellTextRun(ctx, (t, x, y) => ctx.fillText(t, x, y), 'AVA', 50.25, 9, true);
    expect(left).toBe(29);
    expect(fills.map((f) => [f.text, f.x])).toEqual([
      ['A', 29],
      ['V', 36],
      ['A', 43],
    ]);
    expect(ctx.textAlign).toBe('right');
  });
});

const theme = {
  textCell: 11,
  fontUi: 'Arial',
  fontMono: 'monospace',
  accent: '#00f',
  cellErrorFg: '#c00',
  cellBoolFg: '#060',
  cellNumFg: '#000',
  cellAutoFg: '#000',
} as unknown as ResolvedTheme;

const paint = (text: string, format: CellFormat, w: number) => {
  const spy = arialCanvas();
  paintCellText({
    ctx: spy.ctx,
    bounds: { x: 0, y: 0, w, h: 40 },
    theme,
    value: { kind: 'text', value: text },
    formula: null,
    isActive: false,
    isInRange: false,
    format,
    wholePointAdvances: true,
  });
  return spy.fills;
};

describe('cell text on whole-point advances', () => {
  it('paints glyph by glyph at rounded positions', () => {
    const fills = paint('AVAV', {}, 200);
    expect(fills.map((f) => f.text)).toEqual(['A', 'V', 'A', 'V']);
    const xs = fills.map((f) => f.x);
    expect(xs.slice(1).map((x, i) => x - (xs[i] ?? 0))).toEqual([7, 7, 7]);
  });

  it('repeats a fill run by its rounded width', () => {
    // 'fill' is 9 px at 11 pt, so a 24 px wide 4-char column holds two.
    const fills = paint('fill', { align: 'fill' }, 24);
    expect(fills.map((f) => f.text).join('')).toBe('fillfill');
  });

  it('wraps by rounded widths', () => {
    // 'aligned text' is 55 px rounded (56.3 px fractional): it fits a 61 px cell.
    const fills = paint('aligned text', { wrap: true }, 61);
    const ys = new Set(fills.map((f) => f.y));
    expect(ys.size).toBe(1);
  });
});

describe('shrink to fit on whole-point advances', () => {
  it('steps down a whole size until the rounded run fits', () => {
    // A 10-char Mac column holds 55.5 px of text. The scaled size, 3 px,
    // rounds the run to 61 px, so the run paints at 2 px.
    const spy = arialCanvas();
    const sizes: string[] = [];
    const ctx = new Proxy(spy.ctx, {
      set(target, key, value) {
        if (key === 'font') sizes.push(value as string);
        return Reflect.set(target, key, value);
      },
    });
    paintCellText({
      ctx,
      bounds: { x: 0, y: 0, w: 60, h: 40 },
      theme: { ...theme, textCell: (11 * 4) / 3 } as ResolvedTheme,
      value: { kind: 'text', value: 'A very very long shrink to fit sentence' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { shrinkToFit: true },
      zoom: 0.75,
      wholePointAdvances: true,
    });
    expect(sizes.at(-1)).toMatch(/ 2px /);
  });
});

describe('single-line placement on whole-point advances', () => {
  const firstX = (size: number, align: CellFormat['align']): number => {
    const spy = arialCanvas();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 72, h: 40 },
      theme: { ...theme, textCell: size } as ResolvedTheme,
      value: { kind: 'text', value: 'Hxg' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { align },
      wholePointAdvances: true,
    });
    return spy.fills[0]?.x ?? Number.NaN;
  };

  // Excel 'Hxg' ladder, 72 px cells: left origin 1 / 2 / 3 px and the right
  // edge as far in, by size.
  it('keeps a size-dependent inset clear of both edges', () => {
    expect([8, 9, 11, 14, 16, 20].map((size) => firstX(size, 'left'))).toEqual([1, 2, 2, 2, 3, 3]);
    // 'Hxg' advances: 14 px at 8, 20 px at 11, 35 px at 20.
    expect(firstX(8, 'right')).toBe(72 - 1 - 1 - 14);
    expect(firstX(11, 'right')).toBe(72 - 1 - 2 - 20);
    expect(firstX(20, 'right')).toBe(72 - 1 - 3 - 35);
  });

  it('floors a centred run that leaves a half px', () => {
    // 14 px: 'Hxg' is 10 + 7 + 8 = 25 px, centred at 23.5.
    expect(firstX(14, 'center')).toBe(23);
  });
});

describe('rotated text on whole-point advances', () => {
  it('keeps the whole-point inset rather than the zoomed padding', () => {
    // Arial 11 at 200% paints at 22 px: inset 3.25 px, not 3 x 1.5 = 4.5.
    const spy = arialCanvas();
    paintCellText({
      ctx: spy.ctx,
      bounds: { x: 0, y: 0, w: 120, h: 48 },
      theme: { ...theme, textCell: (11 * 4) / 3 } as ResolvedTheme,
      value: { kind: 'text', value: 'Rot' },
      formula: null,
      isActive: false,
      isInRange: false,
      format: { rotation: 45 },
      zoom: 1.5,
      wholePointAdvances: true,
    });
    const ascent = (22 * 10) / 11;
    expect(spy.translates[0]?.[0]).toBeCloseTo(3.25 + ascent * Math.sin(Math.PI / 4));
  });
});
