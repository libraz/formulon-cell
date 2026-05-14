import { describe, expect, it } from 'vitest';

import { paintCheckbox } from '../../../../src/render/painters/controls.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

function makeTheme(over: Partial<ResolvedTheme> = {}): ResolvedTheme {
  return {
    bg: '#ffffff',
    fg: '#1f1f1f',
    rule: '#dddddd',
    ruleStrong: '#888888',
    accent: '#0a8',
    accentFg: '#ffffff',
    headerBg: '#f5f5f5',
    headerFg: '#1f1f1f',
    selBorder: '#0a8',
    selFill: 'rgba(0,170,136,0.1)',
    ...over,
  } as unknown as ResolvedTheme;
}

interface CtxRecording {
  ctx: CanvasRenderingContext2D;
  ops: string[];
  fillStyles: string[];
  rects: [number, number, number, number][];
}

function makeCtxSpy(): CtxRecording {
  const ops: string[] = [];
  const fillStyles: string[] = [];
  const rects: [number, number, number, number][] = [];
  let fillStyle = '';
  let strokeStyle = '';
  const ctx = {
    get fillStyle(): string {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    get strokeStyle(): string {
      return strokeStyle;
    },
    set strokeStyle(v: string) {
      strokeStyle = v;
    },
    set lineWidth(_v: number) {},
    save() {
      ops.push('save');
    },
    restore() {
      ops.push('restore');
    },
    fillRect(x: number, y: number, w: number, h: number) {
      ops.push('fillRect');
      fillStyles.push(fillStyle);
      rects.push([x, y, w, h]);
    },
    strokeRect(x: number, y: number, w: number, h: number) {
      ops.push('strokeRect');
      rects.push([x, y, w, h]);
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
    stroke() {
      ops.push('stroke');
    },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, ops, fillStyles, rects };
}

describe('render/painters/controls — paintCheckbox', () => {
  const bounds = { x: 10, y: 20, w: 30, h: 30 };

  it('returns a 14×14 hit rect centered on the cell', () => {
    const spy = makeCtxSpy();
    const hit = paintCheckbox(spy.ctx, bounds, false, makeTheme());
    expect(hit.rect.w).toBe(14);
    expect(hit.rect.h).toBe(14);
    // center of (10..40, 20..50) is (25, 35); top-left of a 14×14 is (18, 28)
    expect(hit.rect.x).toBe(18);
    expect(hit.rect.y).toBe(28);
  });

  it('checked: fills with accent and strokes white check', () => {
    const spy = makeCtxSpy();
    paintCheckbox(spy.ctx, bounds, true, makeTheme({ accent: '#aabbcc' }));

    expect(spy.fillStyles[0]).toBe('#aabbcc');
    expect(spy.ops).toContain('fillRect');
    expect(spy.ops).toContain('stroke');
    expect(spy.ops[0]).toBe('save');
    expect(spy.ops[spy.ops.length - 1]).toBe('restore');
  });

  it('unchecked: fills with bg and strokes outline only', () => {
    const spy = makeCtxSpy();
    paintCheckbox(spy.ctx, bounds, false, makeTheme({ bg: '#101010' }));

    expect(spy.fillStyles[0]).toBe('#101010');
    expect(spy.ops).toContain('strokeRect');
    expect(spy.ops).not.toContain('stroke'); // no check stroke when unchecked
  });
});
