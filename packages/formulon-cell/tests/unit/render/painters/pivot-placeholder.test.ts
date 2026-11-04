import { describe, expect, it } from 'vitest';
import {
  paintPivotPlaceholder,
  wrapPlaceholderText,
} from '../../../../src/render/painters/pivot-placeholder.js';

describe('wrapPlaceholderText', () => {
  it('breaks at any character to the width', () => {
    expect(wrapPlaceholderText('abcdefg', 3, (s) => s.length)).toEqual(['abc', 'def', 'g']);
  });

  it('drops a space that would start a line', () => {
    expect(wrapPlaceholderText('ab cd', 2, (s) => s.length)).toEqual(['ab', 'cd']);
  });
});

describe('paintPivotPlaceholder (1 pt = 1 px)', () => {
  it('frames the area on its gridlines and places the tag 18 pt in, 20 pt down', () => {
    const calls: { op: string; style: string; args: number[] }[] = [];
    const style = { fill: '', stroke: '' };
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'fillRect' || key === 'strokeRect'
            ? (...args: number[]) =>
                calls.push({
                  op: String(key),
                  style: key === 'fillRect' ? style.fill : style.stroke,
                  args,
                })
            : key === 'measureText'
              ? (s: string) => ({ width: s.length * 10 })
              : () => undefined,
        set: (_t, key, value) => {
          if (key === 'fillStyle') style.fill = value;
          if (key === 'strokeStyle') style.stroke = value;
          return true;
        },
      },
    ) as unknown as CanvasRenderingContext2D;
    // A3:C20 with A 87, B 36, C 30 pt wide and 14 pt rows, at the grid origin (23, 64).
    paintPivotPlaceholder(ctx, {
      box: { x: 23, y: 64, w: 153, h: 252 },
      name: 'PivotTable1',
      prompt: 'prompt',
      ptPx: 1,
      fontFamily: 'sans-serif',
      fontSizePt: 11,
      bg: '#ffffff',
      ink: '#000000',
    });
    expect(calls).toEqual([
      { op: 'fillRect', style: '#ffffff', args: [23, 64, 152, 251] },
      { op: 'strokeRect', style: '#999999', args: [22.5, 63.5, 153, 252] },
      { op: 'fillRect', style: '#ffffff', args: [41, 84, 120, 17] },
      { op: 'strokeRect', style: '#969696', args: [42, 85, 118, 15] },
    ]);
  });
});
