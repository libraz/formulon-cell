import { describe, expect, it } from 'vitest';
import {
  FILTER_BUTTON_SIZE,
  filterButtonRect,
  paintFilterButton,
} from '../../../../src/render/painters/filter-button.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

describe('filterButtonRect', () => {
  it('sits 17x17 one pixel inside the trailing edge and one below the top', () => {
    expect(filterButtonRect({ x: 72, y: 140, w: 72, h: 20 })).toEqual({
      x: 72 + 72 - 1 - 17,
      y: 141,
      w: FILTER_BUTTON_SIZE,
      h: 17,
    });
  });

  it('moves to the leading edge on a right-to-left sheet', () => {
    expect(filterButtonRect({ x: 72, y: 140, w: 72, h: 20 }, true).x).toBe(73);
  });
});

describe('paintFilterButton', () => {
  const run = (state: Parameters<typeof paintFilterButton>[3]) => {
    const strokes: number[][] = [];
    const fills: number[][] = [];
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'strokeRect'
            ? (...a: number[]) => strokes.push(a)
            : key === 'fillRect'
              ? (...a: number[]) => fills.push(a)
              : key === 'createLinearGradient'
                ? () => ({ addColorStop: () => undefined })
                : () => undefined,
        set: () => true,
      },
    ) as unknown as CanvasRenderingContext2D;
    paintFilterButton(
      ctx,
      { x: 10, y: 20, w: 17, h: 17 },
      { accent: '#217346', bg: '#fff' } as ResolvedTheme,
      state,
    );
    return { strokes, fills };
  };

  it('frames the whole button, all four sides', () => {
    const { strokes, fills } = run({});
    expect(fills).toEqual([
      [10, 20, 17, 17],
      [11, 21, 15, 0.5],
    ]);
    expect(strokes).toEqual([[10.5, 20.5, 16, 16]]);
  });
});
