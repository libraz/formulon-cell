import { describe, expect, it } from 'vitest';
import {
  FILTER_BUTTON_SIZE,
  filterButtonRect,
  MAC_PIVOT_BUTTON_LOOK,
  MAC_PIVOT_BUTTON_SIZE,
  paintFilterButton,
  pivotFieldButtonRect,
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

describe('PivotTable field-caption button (Mac)', () => {
  it('is 14x14 flush with the trailing gridline, over the gridline above', () => {
    // Column A of the measured report at 2x: content x 23..90 CSS px, gridline 90..91.
    const rect = pivotFieldButtonRect({ x: 23, y: 55, w: 68, h: 14 }, MAC_PIVOT_BUTTON_SIZE, 1);
    expect(rect).toEqual({ x: 76, y: 54, w: 14, h: 14 });
    // Device px at 2x: 152..179, as measured.
    expect([rect.x * 2, (rect.x + rect.w) * 2 - 1]).toEqual([152, 179]);
  });

  it('draws a two-ring frame at 2x and a 7 px wide arrow', () => {
    const strokes: { color: string; rect: number[] }[] = [];
    const path: number[][] = [];
    let strokeStyle = '';
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'strokeRect'
            ? (...a: number[]) => strokes.push({ color: strokeStyle, rect: a })
            : key === 'moveTo' || key === 'lineTo'
              ? (...a: number[]) => path.push(a)
              : key === 'createLinearGradient'
                ? () => ({ addColorStop: () => undefined })
                : () => undefined,
        set: (_t, key, value) => {
          if (key === 'strokeStyle') strokeStyle = value;
          return true;
        },
      },
    ) as unknown as CanvasRenderingContext2D;
    paintFilterButton(
      ctx,
      { x: 76, y: 54.5, w: 14, h: 14 },
      { accent: '#217346', bg: '#fff' } as ResolvedTheme,
      {},
      MAC_PIVOT_BUTTON_LOOK,
      2,
    );
    expect(strokes).toEqual([
      { color: '#848a90', rect: [76.25, 54.75, 13.5, 13.5] },
      { color: '#93989e', rect: [76.75, 55.25, 12.5, 12.5] },
    ]);
    expect(path).toEqual([
      [79.5, 59.25],
      [86.5, 59.25],
      [83, 63.75],
    ]);
  });
});

describe('filtered PivotTable field button (Mac)', () => {
  it('draws a small arrow at the lower left and a funnel at the right', () => {
    const path: number[][] = [];
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'moveTo' || key === 'lineTo'
            ? (...a: number[]) => path.push(a)
            : key === 'createLinearGradient'
              ? () => ({ addColorStop: () => undefined })
              : () => undefined,
        set: () => true,
      },
    ) as unknown as CanvasRenderingContext2D;
    paintFilterButton(
      ctx,
      { x: 76, y: 54, w: 14, h: 14 },
      { accent: '#217346', bg: '#fff' } as ResolvedTheme,
      { filtered: true },
      MAC_PIVOT_BUTTON_LOOK,
      2,
    );
    // Device px at 2x, as measured: arrow x 156..162 below y 523; funnel bowl x 164..174.
    expect(path.slice(0, 3)).toEqual([
      [78, 61],
      [81.5, 61],
      [79.75, 63.5],
    ]);
    expect(path[3]).toEqual([82, 57]);
    expect(path[4]).toEqual([87.5, 57]);
    expect(path[7]).toEqual([84.25, 64.5]);
  });
});
