import { describe, expect, it } from 'vitest';

import {
  paintFillPreview,
  paintSpillBlocker,
  paintSpillOutline,
} from '../../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

function makeStrokeSpy() {
  const ops: string[] = [];
  const strokes: { style: string; width: number; dash: number[]; alpha: number }[] = [];
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
      strokes.push({ style: strokeStyle, width: lineWidth, dash: [...dash], alpha });
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

describe('render/painters — paintSpillOutline', () => {
  it('strokes a translucent solid ring around the spill range', () => {
    const spy = makeStrokeSpy();
    paintSpillOutline(spy.ctx, { x: 0, y: 0, w: 60, h: 30 }, theme);
    expect(spy.strokes).toHaveLength(1);
    const s = spy.strokes[0];
    expect(s?.style).toBe(theme.accent);
    expect(s?.width).toBe(1);
    expect(s?.dash).toEqual([]);
    // Outline is faded to read as a hint, not a hard border.
    expect(s?.alpha).toBeCloseTo(0.65, 2);
  });
});

describe('render/painters — paintSpillBlocker', () => {
  it('strokes a red dashed outline at 1.5px', () => {
    const spy = makeStrokeSpy();
    paintSpillBlocker(spy.ctx, { x: 0, y: 0, w: 60, h: 30 });
    expect(spy.strokes).toHaveLength(1);
    const s = spy.strokes[0];
    expect(s?.style).toBe('#d83b3b');
    expect(s?.width).toBe(1.5);
    expect(s?.dash).toEqual([3, 3]);
  });
});
