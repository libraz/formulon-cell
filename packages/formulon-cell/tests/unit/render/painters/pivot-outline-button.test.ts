import { describe, expect, it } from 'vitest';
import {
  paintPivotCollapseButton,
  pivotCollapseButtonRect,
  pivotOutlineIndent,
} from '../../../../src/render/painters/pivot-outline-button.js';

describe('PivotTable outline geometry (1 pt = 1 px, as at 100% on the Mac)', () => {
  // Column A of the measured report: content from x 23 CSS px, row 4 from y 269.
  const cell = { x: 23, y: 269, w: 68, h: 14 };

  it('puts the collapse button 4 pt in and 4 pt down, 7 pt square', () => {
    const rect = pivotCollapseButtonRect(cell, 1, 1);
    expect(rect).toEqual({ x: 27, y: 273, w: 7, h: 7 });
    // Device px at 2x: x 54..67, as measured.
    expect([rect.x * 2, (rect.x + rect.w) * 2 - 1]).toEqual([54, 67]);
  });

  it('steps a deeper button one outline step in', () => {
    expect(pivotCollapseButtonRect(cell, 2, 1).x).toBe(38.5);
  });

  it('indents text 11.5 pt per step: outer +14 pt, inner +25.5 pt after the 2.5 pt pad', () => {
    expect(2.5 + pivotOutlineIndent(1, 1)).toBe(14);
    expect(2.5 + pivotOutlineIndent(2, 1)).toBe(25.5);
    expect(2.5 + pivotOutlineIndent(0, 1)).toBe(2.5);
  });

  const recorder = () => {
    const fills: { color: string; rect: number[] }[] = [];
    let fillStyle: unknown = '';
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'fillRect'
            ? (...a: number[]) => fills.push({ color: String(fillStyle), rect: a })
            : key === 'createLinearGradient'
              ? () => ({ addColorStop: () => undefined, toString: () => 'gradient' })
              : () => undefined,
        set: (_t, key, value) => {
          if (key === 'fillStyle') fillStyle = value;
          return true;
        },
      },
    ) as unknown as CanvasRenderingContext2D;
    return { ctx, fills };
  };

  it('paints a two-tone frame with open corners and a 3 x 1 pt minus bar', () => {
    const { ctx, fills } = recorder();
    paintPivotCollapseButton(ctx, { x: 27, y: 273, w: 7, h: 7 }, 1, 2);
    expect(fills).toEqual([
      { color: 'gradient', rect: [27.5, 273.5, 6, 6] },
      { color: '#aab5c9', rect: [27.5, 273, 6, 0.5] },
      { color: '#aab5c9', rect: [27, 273.5, 0.5, 6] },
      { color: '#8b9ebf', rect: [27.5, 279.5, 6, 0.5] },
      { color: '#8b9ebf', rect: [33.5, 273.5, 0.5, 6] },
      { color: '#3d4e7c', rect: [29, 276, 3, 1] },
    ]);
  });

  it('adds the 1 x 3 pt vertical bar of a plus for a collapsed item', () => {
    const { ctx, fills } = recorder();
    paintPivotCollapseButton(ctx, { x: 27, y: 273, w: 7, h: 7 }, 1, 2, true);
    expect(fills.slice(-2)).toEqual([
      { color: '#3d4e7c', rect: [29, 276, 3, 1] },
      { color: '#3d4e7c', rect: [30, 275, 1, 3] },
    ]);
  });
});
