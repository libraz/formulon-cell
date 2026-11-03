import { describe, expect, it } from 'vitest';
import {
  paintActiveCellBox,
  paintRangeOutlineThin,
  paintSelectedFillShade,
  paintSelectionOutline,
} from '../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';

const theme = { bg: '#fefefe', accent: '#217346' } as ResolvedTheme;

const spy = () => {
  const strokes: Array<{ style: string; width: number; rect: number[] }> = [];
  const fills: Array<{ style: string; rect: number[] }> = [];
  let strokeStyle = '';
  let fillStyle = '';
  let lineWidth = 1;
  const ctx = {
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    get lineWidth() {
      return lineWidth;
    },
    set lineWidth(v: number) {
      lineWidth = v;
    },
    save() {},
    restore() {},
    setLineDash() {},
    strokeRect(...rect: number[]) {
      strokes.push({ style: strokeStyle, width: lineWidth, rect });
    },
    fillRect(...rect: number[]) {
      fills.push({ style: fillStyle, rect });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, strokes, fills };
};

describe('multi-range selection', () => {
  it('outlines a range with a one-pixel grey rule on its gridlines and a white inset', () => {
    const s = spy();
    paintRangeOutlineThin(s.ctx, { x: 60, y: 24, w: 120, h: 48 }, theme);
    expect(s.strokes).toEqual([
      { style: '#a5a5a5', width: 1, rect: [59.5, 23.5, 120, 48] },
      { style: '#fefefe', width: 1, rect: [60.5, 24.5, 118, 46] },
    ]);
  });

  it('boxes the active cell one pixel inside its surface', () => {
    const s = spy();
    paintActiveCellBox(s.ctx, { x: 60, y: 24, w: 60, h: 24 }, theme);
    expect(s.strokes).toEqual([{ style: '#217346', width: 1, rect: [61.5, 25.5, 56, 20] }]);
  });
});

describe('selection outline against the grid origin', () => {
  it('keeps the frame inside the grid when no header strip is above it', () => {
    const s = spy();
    paintSelectionOutline(s.ctx, { x: 0, y: 0, w: 60, h: 24 }, theme, false, false, {
      x: 0,
      y: 0,
    });
    expect(s.strokes[0]?.rect).toEqual([1, 1, 59, 23]);
  });
});

describe('paintSelectedFillShade', () => {
  it('darkens a filled cell by 0.776 over the fill rectangle', () => {
    const s = spy();
    paintSelectedFillShade({
      ctx: s.ctx,
      bounds: { x: 10, y: 20, w: 60, h: 30 },
      theme,
      value: { kind: 'blank' },
      formula: null,
      isActive: false,
      isInRange: true,
    });
    expect(s.fills).toEqual([{ style: 'rgba(0, 0, 0, 0.224)', rect: [9, 19, 61, 31] }]);
  });
});
