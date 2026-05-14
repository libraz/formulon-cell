import { describe, expect, it } from 'vitest';

import {
  paintCommentMarker,
  paintLockMarker,
  paintRefHighlight,
  paintTraceArrow,
  paintTraceDot,
  paintValidationChevron,
  TRACE_DEPENDENT_COLOR,
  TRACE_PRECEDENT_COLOR,
} from '../../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

function makeSpy() {
  const ops: string[] = [];
  const fillStyles: string[] = [];
  const strokeStyles: string[] = [];
  const dashes: number[][] = [];
  let fillStyle = '';
  let strokeStyle = '';
  let lineWidth = 1;
  let dash: number[] = [];
  const ctx = {
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
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
    save() {
      ops.push('save');
    },
    restore() {
      ops.push('restore');
    },
    beginPath() {
      ops.push('beginPath');
    },
    moveTo() {
      ops.push('moveTo');
    },
    lineTo() {
      ops.push('lineTo');
    },
    closePath() {
      ops.push('closePath');
    },
    arc() {
      ops.push('arc');
    },
    fill() {
      ops.push('fill');
      fillStyles.push(fillStyle);
    },
    stroke() {
      ops.push('stroke');
      strokeStyles.push(strokeStyle);
    },
    fillRect() {
      ops.push('fillRect');
      fillStyles.push(fillStyle);
    },
    strokeRect() {
      ops.push('strokeRect');
      strokeStyles.push(strokeStyle);
    },
    setLineDash(v: number[]) {
      dash = [...v];
      dashes.push([...v]);
    },
  } as unknown as CanvasRenderingContext2D;
  return {
    ctx,
    ops,
    fillStyles,
    strokeStyles,
    dashes,
    get dash() {
      return dash;
    },
  };
}

const baseTheme = {
  bg: '#fff',
  bgRail: '#eee',
  bgElev: '#fff',
  fg: '#000',
  accent: '#0078d4',
  rule: '#ccc',
} as unknown as ResolvedTheme;

describe('render/painters — trace markers', () => {
  it('paintTraceDot fills a circle at the rect center', () => {
    const spy = makeSpy();
    paintTraceDot(spy.ctx, { x: 10, y: 20, w: 100, h: 30 }, '#abc');
    expect(spy.ops[0]).toBe('save');
    expect(spy.ops).toContain('arc');
    expect(spy.ops).toContain('fill');
    expect(spy.fillStyles).toContain('#abc');
    expect(spy.ops[spy.ops.length - 1]).toBe('restore');
  });

  it('paintTraceArrow exits early when the length is near zero', () => {
    const spy = makeSpy();
    const same = { x: 10, y: 10, w: 20, h: 20 };
    paintTraceArrow(spy.ctx, same, same, '#abc');
    // No save/restore when bailing on degenerate length.
    expect(spy.ops.length).toBe(0);
  });

  it('paintTraceArrow strokes a line then fills the arrow head', () => {
    const spy = makeSpy();
    paintTraceArrow(
      spy.ctx,
      { x: 0, y: 0, w: 10, h: 10 },
      { x: 100, y: 0, w: 10, h: 10 },
      '#1f7ae0',
    );
    expect(spy.ops).toContain('stroke');
    expect(spy.ops).toContain('fill');
    expect(spy.strokeStyles).toContain('#1f7ae0');
    expect(spy.fillStyles).toContain('#1f7ae0');
  });

  it('exposes distinct precedent (blue) and dependent (red) sentinels', () => {
    expect(TRACE_PRECEDENT_COLOR).not.toBe(TRACE_DEPENDENT_COLOR);
    expect(TRACE_PRECEDENT_COLOR).toBe('#1f7ae0');
    expect(TRACE_DEPENDENT_COLOR).toBe('#cf3a4c');
  });
});

describe('render/painters — paintRefHighlight', () => {
  it('strokes a dashed rect inset by 0.5px (crisp 1px line)', () => {
    const spy = makeSpy();
    paintRefHighlight(spy.ctx, { x: 10, y: 20, w: 50, h: 30 }, 0);
    expect(spy.ops).toContain('strokeRect');
    // Last applied dash should be the dashed pattern, not solid.
    expect(spy.dashes.at(-1)).toEqual([5, 3]);
  });

  it('cycles through REF_HIGHLIGHT_COLORS for the index', () => {
    // Two different indexes produce two different stroke styles.
    const a = makeSpy();
    const b = makeSpy();
    paintRefHighlight(a.ctx, { x: 0, y: 0, w: 10, h: 10 }, 0);
    paintRefHighlight(b.ctx, { x: 0, y: 0, w: 10, h: 10 }, 1);
    expect(a.strokeStyles[0]).not.toBe(b.strokeStyles[0]);
  });
});

describe('render/painters — paintValidationChevron', () => {
  it('returns a hit rect at the cell right edge of width 18', () => {
    const spy = makeSpy();
    const hit = paintValidationChevron(spy.ctx, { x: 0, y: 0, w: 100, h: 30 }, baseTheme);
    expect(hit.w).toBe(18);
    expect(hit.x).toBe(100 - 18);
    expect(hit.h).toBeLessThanOrEqual(22);
  });

  it('paints both a filled rail background and a stroked outline', () => {
    const spy = makeSpy();
    paintValidationChevron(spy.ctx, { x: 0, y: 0, w: 100, h: 30 }, baseTheme);
    expect(spy.ops).toContain('fillRect');
    expect(spy.ops).toContain('strokeRect');
    // Chevron triangle path → beginPath, moveTo, lineTo, lineTo, closePath, fill.
    expect(spy.ops).toContain('beginPath');
    expect(spy.ops).toContain('closePath');
  });
});

describe('render/painters — paintLockMarker', () => {
  it('uses theme.accent for both the body fill and the shackle stroke', () => {
    const spy = makeSpy();
    paintLockMarker(spy.ctx, { x: 0, y: 0, w: 40, h: 20 }, baseTheme);
    expect(spy.strokeStyles).toContain(baseTheme.accent);
    expect(spy.fillStyles).toContain(baseTheme.accent);
  });

  it('falls back to a default blue when theme.accent is empty', () => {
    const spy = makeSpy();
    paintLockMarker(spy.ctx, { x: 0, y: 0, w: 40, h: 20 }, { ...baseTheme, accent: '' });
    expect(spy.strokeStyles[0]).toBe('#0078d4');
  });
});

describe('render/painters — paintCommentMarker', () => {
  it('fills a triangle in the upper-right corner using a fixed red', () => {
    const spy = makeSpy();
    paintCommentMarker(spy.ctx, { x: 0, y: 0, w: 60, h: 30 });
    expect(spy.fillStyles).toContain('#d24545');
    expect(spy.ops).toContain('closePath');
    expect(spy.ops).toContain('fill');
  });
});
