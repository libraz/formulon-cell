import { describe, expect, it } from 'vitest';
import { drawConditionalIconArt } from '../../../../src/render/painters/cf-icons.js';

describe('conditional icon fills', () => {
  it('ramps the disc from the Excel highlight to its lower-right shade', () => {
    // Excel red light, centre column (capture space): G holds 161 to about
    // a third of the diagonal, then falls to 118 by two thirds; rim ~(145,69,53).
    const stops: [number, string][] = [];
    const strokes: unknown[] = [];
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'createLinearGradient'
            ? () => ({ addColorStop: (o: number, c: string) => stops.push([o, c]) })
            : () => undefined,
        set: (_t, key, value) => {
          if (key === 'strokeStyle') strokes.push(value);
          return true;
        },
      },
    ) as unknown as CanvasRenderingContext2D;
    drawConditionalIconArt(ctx, 'traffic3', 0);
    expect(stops).toEqual([
      [0.36, '#f69b9c'],
      [0.62, '#f06b6b'],
    ]);
    expect(strokes).toEqual(['#9e3c31']);
  });
});
