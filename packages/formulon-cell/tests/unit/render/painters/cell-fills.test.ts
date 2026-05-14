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
  let fillStyle = '';
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
    stroke() {},
    setLineDash() {},
    set strokeStyle(_v: string) {},
    set lineWidth(_v: number) {},
  } as unknown as CanvasRenderingContext2D;
  return { ctx, rects };
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
});
