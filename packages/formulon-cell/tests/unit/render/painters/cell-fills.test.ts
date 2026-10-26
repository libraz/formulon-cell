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
  let fillStyle = '';
  let strokeStyle = '';
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
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    set lineWidth(_v: number) {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, rects, strokes };
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
    expect(x0).toBeCloseTo(40, 5);
    expect(x1).toBeCloseTo(40, 5);
    expect(y0).toBeCloseTo(20, 5);
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
    expect(cx).toBeCloseTo(40, 5);
    expect(cy).toBeCloseTo(35, 5);
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
