import { describe, expect, it, vi } from 'vitest';
import { paintDataBar } from '../../../../src/render/grid/data-bar.js';

describe('data-bar painting', () => {
  it.each([
    { axis: 0, direction: 'right' as const, start: 10, end: 50, x: 10 },
    { axis: 0.5, direction: 'left' as const, start: 60, end: 20, x: 20 },
    { axis: 1, direction: 'left' as const, start: 110, end: 70, x: 70 },
  ])(
    'fades from axis $axis toward the $direction endpoint',
    ({ axis, direction, start, end, x }) => {
      const addColorStop = vi.fn();
      const createLinearGradient = vi.fn(() => ({ addColorStop }));
      const fillRect = vi.fn();
      const ctx = {
        save: vi.fn(),
        restore: vi.fn(),
        createLinearGradient,
        fillRect,
      } as unknown as CanvasRenderingContext2D;

      paintDataBar(
        ctx,
        { x: 10, y: 20, w: 100, h: 30 },
        {
          bar: 0.4,
          barAxis: axis,
          barDirection: direction,
          barColor: '#0078d4',
          barGradient: true,
        },
      );

      expect(createLinearGradient).toHaveBeenCalledWith(start, 0, end, 0);
      expect(addColorStop.mock.calls).toEqual([
        [0, '#0078d4'],
        [1, 'rgba(255,255,255,0.25)'],
      ]);
      expect(fillRect).toHaveBeenCalledWith(x, 21, 40, 28);
    },
  );
});
