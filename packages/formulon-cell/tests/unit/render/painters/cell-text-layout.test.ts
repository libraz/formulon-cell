import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../../../src/engine/types.js';
import {
  fontMetricsBox,
  paintCellText,
  shrinkFontSizeToFit,
} from '../../../../src/render/painters/text.js';
import type { CellFormat } from '../../../../src/store/store.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

/** Canvas spy with Arial-like font-box metrics (10/2 at 11 px) reported from
 *  the current `textBaseline`, and half-em advances per character. */
function makeSpy() {
  const fills: Array<{ text: string; x: number; y: number }> = [];
  const lines: Array<{ x0: number; x1: number; y: number }> = [];
  const transforms: Array<['translate', number, number] | ['rotate', number]> = [];
  let font = '400 11px Arial';
  let moveTo: [number, number] = [0, 0];
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
      transforms.push(['translate', x, y]);
    },
    rotate(a: number): void {
      transforms.push(['rotate', a]);
    },
    moveTo(x: number, y: number): void {
      moveTo = [x, y];
    },
    lineTo(x: number): void {
      lines.push({ x0: moveTo[0], x1: x, y: moveTo[1] });
    },
    stroke(): void {},
    measureText(text: string): TextMetrics {
      const s = size();
      const shift = this.textBaseline === 'middle' ? -3 : 0;
      return {
        width: [...text].length * s * 0.5,
        fontBoundingBoxAscent: (s * 10) / 11 + shift,
        fontBoundingBoxDescent: (s * 2) / 11 - shift,
        actualBoundingBoxAscent: s * 0.72,
        actualBoundingBoxDescent: s * 0.22,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number): void {
      fills.push({ text, x, y });
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills, lines, transforms };
}

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

const bounds = { x: 0, y: 0, w: 60, h: 48 };

const paint = (
  value: CellValue,
  format: CellFormat,
  over: { w?: number; zoom?: number; textCell?: number } = {},
) => {
  const spy = makeSpy();
  paintCellText({
    ctx: spy.ctx,
    bounds: { ...bounds, w: over.w ?? bounds.w },
    theme: { ...theme, textCell: over.textCell ?? 11 } as ResolvedTheme,
    value,
    formula: null,
    isActive: false,
    isInRange: false,
    format,
    zoom: over.zoom,
  });
  return spy;
};

const text = (value: string): CellValue => ({ kind: 'text', value });

describe('fontMetricsBox', () => {
  it('measures from the alphabetic baseline whatever baseline is set', () => {
    const { ctx } = makeSpy();
    ctx.textBaseline = 'middle';
    expect(fontMetricsBox(ctx, 11)).toEqual({ ascent: 10, descent: 2 });
    expect(ctx.textBaseline).toBe('middle');
  });

  it('keeps wrapped text on the bottom baseline after a painter left middle set', () => {
    const spy = makeSpy();
    spy.ctx.textBaseline = 'middle';
    paintCellText({
      ctx: spy.ctx,
      bounds,
      theme,
      value: text('one'),
      formula: null,
      isActive: false,
      isInRange: false,
      format: { wrap: true },
    });
    expect(spy.fills[0]?.y).toBe(45);
  });
});

describe('rotated cell text', () => {
  // 'Rot90' is 27.5 px wide; padX is 3 and the font box is 10 over 2.
  const placed = (rotation: number, format: CellFormat = {}) => {
    const spy = paint(text('Rot90'), { rotation, ...format });
    const [translate, rotate] = spy.transforms;
    return { x: translate?.[1], y: translate?.[2], angle: rotate?.[1] };
  };

  it('rotates counterclockwise for positive angles', () => {
    expect(placed(90).angle).toBeCloseTo(-Math.PI / 2);
    expect(placed(-45).angle).toBeCloseTo(Math.PI / 4);
  });

  it('puts a rising run at the bottom-left by its ascent corner', () => {
    const p = placed(45);
    expect(p.x).toBeCloseTo(3 + 10 * Math.SQRT1_2);
    expect(p.y).toBe(45);
  });

  it('puts +90 at the right, keeping its descent one padding clear', () => {
    const p = placed(90);
    expect(p.x).toBeCloseTo(60 - 0.5 - 3 - 3 - 2);
    expect(p.y).toBe(45);
  });

  it('puts -90 at the left with its end on the bottom baseline', () => {
    const p = placed(-90);
    expect(p.x).toBeCloseTo(3 + 2);
    expect(p.y).toBeCloseTo(45 - 27.5);
  });

  it('puts a falling run at the right with its lower end on the baseline', () => {
    const p = placed(-45);
    const maxX = 27.5 * Math.SQRT1_2 + 10 * Math.SQRT1_2;
    expect(p.x).toBeCloseTo(53.5 - maxX);
    expect(p.y).toBeCloseTo(45 - 27.5 * Math.SQRT1_2);
  });

  it('honours explicit centre and top alignment by the line box', () => {
    const p = placed(90, { align: 'center', vAlign: 'top' });
    expect(p.x).toBeCloseTo(30 - (-10 + 2) / 2);
    expect(p.y).toBeCloseTo(1 + 27.5);
  });
});

describe('fill alignment', () => {
  it('repeats the run as many whole times as the cell holds', () => {
    expect(paint(text('ab'), { align: 'fill' }).fills[0]?.text).toBe('abababab');
  });

  it('paints a run wider than half the cell once', () => {
    expect(paint(text('fill'), { align: 'fill' }, { w: 24 }).fills[0]?.text).toBe('fill');
  });
});

describe('shrink to fit', () => {
  it('has no minimum size beyond one pixel', () => {
    expect(shrinkFontSizeToFit(11, 96, 55.5)).toBe(6);
    expect(shrinkFontSizeToFit(11, 400, 140)).toBe(3);
    expect(shrinkFontSizeToFit(11, 4000, 30)).toBe(1);
  });

  it('keeps the full-size baseline for a shrunk bottom-aligned run', () => {
    const full = paint(text('x'), {});
    const shrunk = paint(text('Shrink this long text'), { shrinkToFit: true });
    expect(shrunk.fills[0]?.y).toBe(full.fills[0]?.y);
  });
});

describe('underline rules', () => {
  // Baseline 45 at 11 px puts the single rule on the 46.5 stroke row.
  const rows = (format: CellFormat, value: CellValue = text('Hxg')) =>
    paint(value, format).lines.map((l) => l.y);

  it('brackets the single rule for double', () => {
    expect(rows({ underline: true })).toEqual([46.5]);
    expect(rows({ underline: 'double' })).toEqual([44.5, 46.5]);
  });

  it('drops the accounting styles one pixel lower', () => {
    expect(rows({ underline: 'singleAccounting' })).toEqual([47.5]);
    expect(rows({ underline: 'doubleAccounting' })).toEqual([47.5, 49.5]);
  });

  it('spans the padded cell for text and the run for numbers', () => {
    const textRule = paint(text('Hxg'), { underline: 'singleAccounting' }).lines[0];
    expect([textRule?.x0, textRule?.x1]).toEqual([3, 57]);
    const numberRule = paint({ kind: 'number', value: 1234 }, { underline: 'singleAccounting' })
      .lines[0];
    expect([numberRule?.x0, numberRule?.x1]).toEqual([35, 57]);
  });
});

describe('wrapped text', () => {
  it('rounds the line pitch in rendered pixels', () => {
    // An 11 pt run on the 0.75 profile renders at 11 px: 14 px lines, not 14.25.
    const spy = paint(
      text('Line one\nLine two'),
      { wrap: true },
      { zoom: 0.75, textCell: 44 / 3, w: 80 },
    );
    expect(spy.fills).toHaveLength(2);
    expect((spy.fills[1]?.y ?? 0) - (spy.fills[0]?.y ?? 0)).toBe(14);
  });
});
