import { describe, expect, it } from 'vitest';

import {
  paintFillPreview,
  paintSpillBlocker,
  paintSpillOutline,
  paintSpillTarget,
} from '../../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

function makeStrokeSpy() {
  const ops: string[] = [];
  const strokes: {
    style: string;
    width: number;
    dash: number[];
    alpha: number;
    rect: number[];
    shadow: { color: string; blur: number };
  }[] = [];
  let shadowColor = '';
  let shadowBlur = 0;
  let strokeStyle = '';
  let lineWidth = 1;
  let dash: number[] = [];
  let alpha = 1;
  const ctx = {
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    get lineWidth() {
      return lineWidth;
    },
    set lineWidth(v: number) {
      lineWidth = v;
    },
    get shadowColor() {
      return shadowColor;
    },
    set shadowColor(v: string) {
      shadowColor = v;
    },
    get shadowBlur() {
      return shadowBlur;
    },
    set shadowBlur(v: number) {
      shadowBlur = v;
    },
    get globalAlpha() {
      return alpha;
    },
    set globalAlpha(v: number) {
      alpha = v;
    },
    save() {
      ops.push('save');
    },
    restore() {
      ops.push('restore');
    },
    setLineDash(v: number[]) {
      dash = [...v];
    },
    strokeRect(x: number, y: number, w: number, h: number) {
      ops.push(`strokeRect:${x},${y},${w},${h}`);
      strokes.push({
        style: strokeStyle,
        width: lineWidth,
        dash: [...dash],
        alpha,
        rect: [x, y, w, h],
        shadow: { color: shadowColor, blur: shadowBlur },
      });
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, ops, strokes };
}

const theme = {
  accent: '#0078d4',
  rule: '#ccc',
  fg: '#000',
  bg: '#fff',
} as unknown as ResolvedTheme;

describe('render/painters — paintFillPreview', () => {
  it('strokes a dashed accent-coloured rectangle inset by 0.5px', () => {
    const spy = makeStrokeSpy();
    paintFillPreview(spy.ctx, { x: 10, y: 20, w: 100, h: 40 }, theme);
    expect(spy.strokes).toHaveLength(1);
    const s = spy.strokes[0];
    expect(s?.style).toBe(theme.accent);
    expect(s?.dash).toEqual([4, 3]);
    expect(s?.width).toBe(1.5);
    expect(spy.ops[0]).toBe('save');
    expect(spy.ops[spy.ops.length - 1]).toBe('restore');
  });
});

// Mac desktop, measured: a 1 pt accent (sRGB 49,102,194 / 48,101,192) on the
// range's outer gridlines with a soft outer shadow; the blocked box dashes 14 / 3 pt.
describe('render/painters — paintSpillOutline', () => {
  it('strokes a solid 1 pt accent ring over the range gridlines, with a soft shadow', () => {
    const spy = makeStrokeSpy();
    paintSpillOutline(spy.ctx, { x: 10, y: 20, w: 60, h: 30 }, theme, 1);
    expect(spy.strokes).toHaveLength(1);
    const s = spy.strokes[0];
    expect(s?.style).toBe('#3065c0');
    expect(s?.width).toBe(1);
    expect(s?.dash).toEqual([]);
    expect(s?.alpha).toBe(1);
    // The leading gridlines sit just before the range; the trailing ones are its own.
    expect(s?.rect).toEqual([9.5, 19.5, 60, 30]);
    expect(s?.shadow.blur).toBeGreaterThan(0);
    expect(s?.shadow.color).not.toBe('');
  });

  it('scales the stroke with the pt-to-px factor', () => {
    const spy = makeStrokeSpy();
    paintSpillOutline(spy.ctx, { x: 0, y: 0, w: 60, h: 30 }, theme, 2);
    expect(spy.strokes[0]?.width).toBe(2);
  });
});

describe('render/painters — paintSpillTarget', () => {
  it('dashes the would-be spill range 14 pt on, 3 pt off, in the same accent', () => {
    const spy = makeStrokeSpy();
    paintSpillTarget(spy.ctx, { x: 10, y: 20, w: 60, h: 100 }, 1);
    const s = spy.strokes[0];
    expect(s?.style).toBe('#3065c0');
    expect(s?.width).toBe(1);
    expect(s?.dash).toEqual([14, 3]);
    expect(s?.rect).toEqual([9.5, 19.5, 60, 100]);
    expect(s?.shadow.blur).toBeGreaterThan(0);
    const scaled = makeStrokeSpy();
    paintSpillTarget(scaled.ctx, { x: 0, y: 0, w: 60, h: 100 }, 2);
    expect(scaled.strokes[0]?.dash).toEqual([28, 6]);
  });
});

describe('render/painters — spill boxes off the Mac profile', () => {
  it('keeps the faint 1px accent ring, the 4 / 3 dashed target and the red blocker', () => {
    const ring = makeStrokeSpy();
    paintSpillOutline(ring.ctx, { x: 0, y: 0, w: 60, h: 30 }, theme);
    expect(ring.strokes[0]).toMatchObject({
      style: theme.accent,
      width: 1,
      dash: [],
      rect: [0.5, 0.5, 59, 29],
    });
    expect(ring.strokes[0]?.alpha).toBeCloseTo(0.65, 2);
    expect(ring.strokes[0]?.shadow.blur).toBe(0);
    const target = makeStrokeSpy();
    paintSpillTarget(target.ctx, { x: 0, y: 0, w: 60, h: 30 });
    expect(target.strokes[0]).toMatchObject({ width: 1, dash: [4, 3], rect: [0.5, 0.5, 59, 29] });
    const blocker = makeStrokeSpy();
    paintSpillBlocker(blocker.ctx, { x: 0, y: 0, w: 60, h: 30 });
    expect(blocker.strokes[0]).toMatchObject({ style: '#d83b3b', width: 1.5, dash: [3, 3] });
  });
});
