import { describe, expect, it } from 'vitest';
import { drawConditionalIconArt } from '../../../../src/render/painters/cf-icons.js';

describe('conditional icon fills', () => {
  it('spans the light-to-dark ramp over the middle of the disc diagonal', () => {
    const stops: [number, string][] = [];
    const ctx = new Proxy(
      {},
      {
        get: (_t, key) =>
          key === 'createLinearGradient'
            ? () => ({ addColorStop: (o: number, c: string) => stops.push([o, c]) })
            : () => undefined,
        set: () => true,
      },
    ) as unknown as CanvasRenderingContext2D;
    drawConditionalIconArt(ctx, 'traffic3', 0);
    expect(stops).toEqual([
      [0.25, '#f7a3a2'],
      [0.75, '#ec6d69'],
    ]);
  });
});
