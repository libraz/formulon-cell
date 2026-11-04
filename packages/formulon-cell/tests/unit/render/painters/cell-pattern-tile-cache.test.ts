import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CellPaintCtx } from '../../../../src/render/painters/types.js';
import { paintCellFill } from '../../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

const noop = (): void => {};
const ctx = {
  fillStyle: '',
  fillRect: noop,
  save: noop,
  restore: noop,
  beginPath: noop,
  rect: noop,
  fill: noop,
} as unknown as CanvasRenderingContext2D;

describe('paintCellFill pattern tiles', () => {
  afterEach(() => vi.restoreAllMocks());

  it("derives a tile's runs once across cells", () => {
    const map = vi.spyOn(Array.prototype, 'map');
    for (let i = 0; i < 5; i += 1) {
      paintCellFill({
        ctx,
        theme: { bg: '#fff' } as ResolvedTheme,
        bounds: { x: i * 10, y: 0, w: 10, h: 10 },
        format: { fillPattern: 'lightGrid' },
        dpr: 1,
      } as unknown as CellPaintCtx);
    }
    const tileMaps = map.mock.contexts.filter(
      (c) => Array.isArray(c) && typeof c[0] === 'string' && /^[X.]{8}$/.test(c[0]),
    );
    expect(tileMaps.length).toBeLessThanOrEqual(1);
  });
});
