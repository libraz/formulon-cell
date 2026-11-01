import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../../../src/engine/types.js';
import {
  type CellPaintCtx,
  CONDITIONAL_ICON_GUTTER,
  paintCellBackground,
  paintCellFill,
  paintConditionalIcon,
} from '../../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

function makeFillSpy() {
  const rects: { style: string; rect: [number, number, number, number] }[] = [];
  const strokes: string[] = [];
  const transforms: { op: 'translate' | 'scale'; x: number; y: number }[] = [];
  const stops: string[] = [];
  let fillStyle = '';
  let strokeStyle = '';
  const gradient = () => ({
    addColorStop(_offset: number, color: string) {
      stops.push(color);
    },
  });
  const ctx = {
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    fillRect(x: number, y: number, w: number, h: number) {
      rects.push({ style: fillStyle, rect: [x, y, w, h] });
    },
    save() {},
    restore() {},
    beginPath() {},
    moveTo() {},
    lineTo() {},
    closePath() {},
    rect() {},
    clip() {},
    arc() {},
    fill() {
      rects.push({ style: fillStyle, rect: [0, 0, 0, 0] });
    },
    stroke() {
      strokes.push(strokeStyle);
    },
    setLineDash() {},
    translate(x: number, y: number) {
      transforms.push({ op: 'translate', x, y });
    },
    scale(x: number, y: number) {
      transforms.push({ op: 'scale', x, y });
    },
    createLinearGradient: gradient,
    createRadialGradient: gradient,
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    set lineWidth(_v: number) {},
    set lineCap(_v: string) {},
    set lineJoin(_v: string) {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, rects, strokes, transforms, stops };
}

const theme = {
  bg: '#fff',
  bgRail: '#eee',
  bgElev: '#fafafa',
  bgHeader: '#eee',
  fg: '#000',
  accent: '#0078d4',
  accentSoft: 'rgba(0,120,212,0.1)',
  rule: '#ccc',
} as unknown as ResolvedTheme;

function makeCtx(over: Partial<CellPaintCtx>): CellPaintCtx {
  const spy = makeFillSpy();
  return {
    ctx: spy.ctx,
    theme,
    bounds: { x: 10, y: 20, w: 60, h: 30 },
    value: { kind: 'blank' } as CellValue,
    formula: null,
    isActive: false,
    isInRange: false,
    ...over,
  };
}

describe('render/painters — paintCellBackground', () => {
  it('paints theme.bgElev for the active cell', () => {
    const spy = makeFillSpy();
    paintCellBackground({ ...makeCtx({ isActive: true }), ctx: spy.ctx });
    expect(spy.rects[0]?.style).toBe(theme.bgElev);
  });

  it('clears only the active surface, keeping the trailing rules shaded', () => {
    const spy = makeFillSpy();
    paintCellBackground({ ...makeCtx({ isActive: true, dpr: 2 }), ctx: spy.ctx });
    expect(spy.rects[0]?.rect).toEqual([10, 20, 59, 29]);
    const rtl = makeFillSpy();
    paintCellBackground({ ...makeCtx({ isActive: true, dpr: 2, rtl: true }), ctx: rtl.ctx });
    expect(rtl.rects[0]?.rect).toEqual([11, 20, 59, 29]);
  });

  it('paints accentSoft for in-range cells', () => {
    const spy = makeFillSpy();
    paintCellBackground({ ...makeCtx({ isInRange: true }), ctx: spy.ctx });
    expect(spy.rects[0]?.style).toBe(theme.accentSoft);
  });

  it('no-ops for non-active, non-in-range cells (global bg paints them)', () => {
    const spy = makeFillSpy();
    paintCellBackground({ ...makeCtx({}), ctx: spy.ctx });
    expect(spy.rects).toHaveLength(0);
  });

  it('active wins over in-range when both are true', () => {
    const spy = makeFillSpy();
    paintCellBackground({ ...makeCtx({ isActive: true, isInRange: true }), ctx: spy.ctx });
    expect(spy.rects).toHaveLength(1);
    expect(spy.rects[0]?.style).toBe(theme.bgElev);
  });
});

describe('render/painters — paintCellFill', () => {
  it('paints the user-set fill when format.fill is set', () => {
    const spy = makeFillSpy();
    paintCellFill({ ...makeCtx({ format: { fill: '#ffe0b0' } }), ctx: spy.ctx });
    expect(spy.rects[0]?.style).toBe('#ffe0b0');
  });

  it('paints fill patterns over the background fill', () => {
    const spy = makeFillSpy();
    paintCellFill({
      ...makeCtx({
        format: { fill: '#ffeecc', fillPattern: 'horizontal', fillPatternColor: '#336699' },
      }),
      ctx: spy.ctx,
    });
    expect(spy.rects[0]?.style).toBe('#ffeecc');
    expect(spy.strokes).toContain('#336699');
  });

  it('no-ops when format.fill is missing', () => {
    const spy = makeFillSpy();
    paintCellFill({ ...makeCtx({ format: {} }), ctx: spy.ctx });
    expect(spy.rects).toHaveLength(0);
  });
});

describe('render/painters — paintConditionalIcon', () => {
  it('exposes a 16px gutter so paintCellText can right-shift around the glyph', () => {
    expect(CONDITIONAL_ICON_GUTTER).toBe(16);
  });

  it('arrows3 paints a stroke + filled head triangle for the slot', () => {
    const spy = makeFillSpy();
    paintConditionalIcon(spy.ctx, { x: 0, y: 0, w: 20, h: 20 }, 'arrows3', 1);
    // arrow body + arrow head — at least one fill call lands.
    expect(spy.rects.length).toBeGreaterThan(0);
  });

  it('traffic3 fills a single circle for the slot', () => {
    const spy = makeFillSpy();
    paintConditionalIcon(spy.ctx, { x: 0, y: 0, w: 20, h: 20 }, 'traffic3', 2);
    expect(spy.rects.length).toBeGreaterThan(0);
  });

  it('stars3 differentiates full / half / empty by paint count', () => {
    const full = makeFillSpy();
    const empty = makeFillSpy();
    paintConditionalIcon(full.ctx, { x: 0, y: 0, w: 20, h: 20 }, 'stars3', 2);
    paintConditionalIcon(empty.ctx, { x: 0, y: 0, w: 20, h: 20 }, 'stars3', 0);
    // Full has at least one fill; empty has only the outline stroke (no fill).
    expect(full.rects.length).toBeGreaterThanOrEqual(1);
    expect(empty.rects.length).toBe(0);
  });

  it('clamps out-of-range slot indices', () => {
    const spy = makeFillSpy();
    // Should not throw on slot = 99 (traffic3 has only 3 slots).
    expect(() =>
      paintConditionalIcon(spy.ctx, { x: 0, y: 0, w: 20, h: 20 }, 'traffic3', 99),
    ).not.toThrow();
  });

  it('places the 11px icon box 1px in and follows the vertical alignment', () => {
    const origin = (vAlign?: 'top' | 'middle' | 'bottom', zoom?: number, rtl = false) => {
      const spy = makeFillSpy();
      const rect = { x: 10, y: 100, w: 48, h: 40 };
      paintConditionalIcon(spy.ctx, rect, 'traffic3', 0, rtl, vAlign, zoom);
      return spy.transforms;
    };
    // Bottom (default): 2px clear of the 1px gridline the cell owns.
    expect(origin()[0]).toEqual({ op: 'translate', x: 11, y: 100 + 40 - 1 - 2 - 11 });
    expect(origin('bottom')[0]).toEqual(origin()[0]);
    expect(origin('top')[0]).toEqual({ op: 'translate', x: 11, y: 100 });
    expect(origin('middle')[0]).toEqual({ op: 'translate', x: 11, y: 100 + (40 - 1 - 2 - 11) / 2 });
    expect(origin(undefined, undefined, true)[0]).toEqual({
      op: 'translate',
      x: 10 + 48 - 12,
      y: 126,
    });
    expect(origin('top', 2)).toEqual([
      { op: 'translate', x: 12, y: 100 },
      { op: 'scale', x: 2, y: 2 },
    ]);
  });

  it('shades 4/5-arrow middle slots yellow and the ends red / green', () => {
    const tones = (kind: 'arrows4' | 'arrows5', slot: number) => {
      const spy = makeFillSpy();
      paintConditionalIcon(spy.ctx, { x: 0, y: 0, w: 20, h: 20 }, kind, slot);
      return spy.stops[0];
    };
    expect(tones('arrows5', 0)).toBe(tones('arrows4', 0));
    expect(tones('arrows5', 4)).toBe(tones('arrows4', 3));
    for (const slot of [1, 2, 3]) expect(tones('arrows5', slot)).toBe(tones('arrows4', 1));
    expect(tones('arrows4', 2)).toBe(tones('arrows4', 1));
    expect(tones('arrows4', 1)).not.toBe(tones('arrows4', 0));
    expect(tones('arrows4', 1)).not.toBe(tones('arrows4', 3));
  });

  it('lights one rating bar per level, starting from one (4) or none (5)', () => {
    const lit = (kind: 'ratings4' | 'ratings5', slot: number) => {
      const spy = makeFillSpy();
      paintConditionalIcon(spy.ctx, { x: 0, y: 0, w: 20, h: 20 }, kind, slot);
      return spy.rects.filter((r) => r.style === '#153a6a').length;
    };
    expect([0, 1, 2, 3].map((slot) => lit('ratings4', slot))).toEqual([1, 2, 3, 4]);
    expect([0, 1, 2, 3, 4].map((slot) => lit('ratings5', slot))).toEqual([0, 1, 2, 3, 4]);
  });

  it('paints expanded icon families without external assets', () => {
    for (const kind of [
      'triangles3',
      'symbols3',
      'flags3',
      'quarters5',
      'bars5',
      'boxes5',
    ] as const) {
      const spy = makeFillSpy();
      paintConditionalIcon(spy.ctx, { x: 0, y: 0, w: 20, h: 20 }, kind, 1);
      expect(spy.rects.length).toBeGreaterThan(0);
    }
  });
});

describe('render/painters — gradient fills', () => {
  const gradientSpy = () => {
    const created: Array<{ kind: 'linear' | 'radial'; args: number[]; stops: [number, string][] }> =
      [];
    const rects: Array<{ style: unknown }> = [];
    let fillStyle: unknown = '';
    const make =
      (kind: 'linear' | 'radial') =>
      (...args: number[]) => {
        const entry = { kind, args, stops: [] as [number, string][] };
        created.push(entry);
        return {
          addColorStop(offset: number, color: string) {
            entry.stops.push([offset, color]);
          },
          entry,
        };
      };
    const ctx = {
      get fillStyle() {
        return fillStyle;
      },
      set fillStyle(v: unknown) {
        fillStyle = v;
      },
      fillRect() {
        rects.push({ style: fillStyle });
      },
      save() {},
      restore() {},
      createLinearGradient: make('linear'),
      createRadialGradient: make('radial'),
    } as unknown as CanvasRenderingContext2D;
    return { ctx, created, rects };
  };
  const stops = [
    { position: 0, color: '#ff0000' },
    { position: 1, color: '#0000ff' },
  ];

  it('paints a 90 degree linear gradient top to bottom across the cell', () => {
    const spy = gradientSpy();
    paintCellFill({
      ...makeCtx({}),
      ctx: spy.ctx,
      format: {
        fillGradient: { type: 'linear', degree: 90, left: 0, right: 0, top: 0, bottom: 0, stops },
      },
    });
    expect(spy.created).toHaveLength(1);
    const [x0, y0, x1, y1] = spy.created[0]?.args ?? [];
    // The fill rect is {39,19,81,31}: the cell grown over the gridline before it.
    expect(x0).toBeCloseTo(39.5, 5);
    expect(x1).toBeCloseTo(39.5, 5);
    expect(y0).toBeCloseTo(19, 5);
    expect(y1).toBeCloseTo(50, 5);
    expect(spy.created[0]?.stops).toEqual([
      [0, '#ff0000'],
      [1, '#0000ff'],
    ]);
    expect(spy.rects).toHaveLength(1);
  });

  it('paints a path gradient outward from its inner rectangle', () => {
    const spy = gradientSpy();
    paintCellFill({
      ...makeCtx({}),
      ctx: spy.ctx,
      format: {
        fillGradient: {
          type: 'path',
          degree: 0,
          left: 0.5,
          right: 0.5,
          top: 0.5,
          bottom: 0.5,
          stops,
        },
      },
    });
    expect(spy.created[0]?.kind).toBe('radial');
    const [cx, cy] = spy.created[0]?.args ?? [];
    expect(cx).toBeCloseTo(39.5, 5);
    expect(cy).toBeCloseTo(34.5, 5);
  });

  it('a gradient replaces the solid fill colour', () => {
    const spy = gradientSpy();
    paintCellFill({
      ...makeCtx({}),
      ctx: spy.ctx,
      format: {
        fill: '#00ff00',
        fillGradient: { type: 'linear', degree: 0, left: 0, right: 0, top: 0, bottom: 0, stops },
      },
    });
    expect(spy.rects).toHaveLength(1);
    expect(spy.rects[0]?.style).not.toBe('#00ff00');
  });
});
