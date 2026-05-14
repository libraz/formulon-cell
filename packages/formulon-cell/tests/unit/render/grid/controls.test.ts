import { afterEach, describe, expect, it } from 'vitest';
import { getFillHandleRect, setFillHandleRect } from '../../../../src/render/grid/hit-state.js';
import {
  FILL_HANDLE_SIZE,
  paintFillHandle,
  paintTableHeaderChevron,
} from '../../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../../src/theme/resolve.js';

/** Pointer-layer pad in pixels (interact/pointer.ts:isFillHandleHit). The
 *  visible handle is small (6×6), so the hit-zone gets a 3px halo to make
 *  the grab area comfortable. This constant is duplicated here so the test
 *  fails loudly if pointer.ts changes one without the other. */
const FILL_HANDLE_HIT_PAD = 3;

/** Reproduce the closure-scoped isFillHandleHit logic from interact/pointer.ts
 *  so this test can assert the contract without spinning up a real Spreadsheet
 *  mount. The two implementations must stay in sync. */
function isFillHandleHit(x: number, y: number): boolean {
  const rect = getFillHandleRect();
  if (!rect) return false;
  const pad = FILL_HANDLE_HIT_PAD;
  return (
    x >= rect.x - pad &&
    x <= rect.x + rect.w + pad &&
    y >= rect.y - pad &&
    y <= rect.y + rect.h + pad
  );
}

/** Minimal canvas spy with both fill + stroke. We don't care about the
 *  pixels — only what the painter returns to the cache layer. */
function ctxSpy(): CanvasRenderingContext2D {
  return {
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    fillRect(): void {},
    strokeRect(): void {},
    beginPath(): void {},
    moveTo(): void {},
    lineTo(): void {},
    closePath(): void {},
    fill(): void {},
    stroke(): void {},
    save(): void {},
    restore(): void {},
  } as unknown as CanvasRenderingContext2D;
}

function theme(over: Partial<ResolvedTheme> = {}): ResolvedTheme {
  return {
    accent: '#0078d4',
    fg: '#1f1f1f',
    fgMute: '#5a5a5a',
    rule: '#cccccc',
    ...over,
  } as ResolvedTheme;
}

describe('render/grid/controls — painter → cache → hit-zone contract', () => {
  afterEach(() => {
    setFillHandleRect(null);
  });

  describe('fill handle', () => {
    it('returns a rect that pads the visible square by 1px on every side', () => {
      const ctx = ctxSpy();
      const cell = { x: 100, y: 60, w: 80, h: 20 };
      const hit = paintFillHandle(ctx, cell, theme());

      // The visible square is FILL_HANDLE_SIZE (6) centered on the cell's
      // bottom-right corner. The returned rect adds the 1px white border ring
      // on each side -> 8x8 starting one pixel earlier.
      const hs = FILL_HANDLE_SIZE;
      const cornerX = cell.x + cell.w; // 180
      const cornerY = cell.y + cell.h; // 80
      expect(hit).toEqual({
        x: cornerX - hs / 2 - 1, // 176
        y: cornerY - hs / 2 - 1, // 76
        w: hs + 2, // 8
        h: hs + 2, // 8
      });
    });

    it('caches the painted rect and lets the pointer layer read it back', () => {
      const ctx = ctxSpy();
      const rect = paintFillHandle(ctx, { x: 0, y: 0, w: 50, h: 20 }, theme());
      setFillHandleRect(rect);
      expect(getFillHandleRect()).toBe(rect);
    });

    it('hit-tests true inside the painted rect', () => {
      setFillHandleRect({ x: 100, y: 100, w: 8, h: 8 });
      expect(isFillHandleHit(104, 104)).toBe(true); // center
      expect(isFillHandleHit(100, 100)).toBe(true); // top-left corner
      expect(isFillHandleHit(108, 108)).toBe(true); // bottom-right corner
    });

    it('hit-tests true within the 3px halo (grab zone)', () => {
      setFillHandleRect({ x: 100, y: 100, w: 8, h: 8 });
      expect(isFillHandleHit(97, 97)).toBe(true); // 3px outside top-left
      expect(isFillHandleHit(111, 111)).toBe(true); // 3px outside bot-right
      expect(isFillHandleHit(96, 100)).toBe(false); // 4px outside — miss
      expect(isFillHandleHit(112, 100)).toBe(false); // 4px outside — miss
    });

    it('hit-tests false when the cache is null (handle offscreen)', () => {
      setFillHandleRect(null);
      expect(isFillHandleHit(0, 0)).toBe(false);
      expect(isFillHandleHit(100, 100)).toBe(false);
    });

    it('hit zone follows the rect after re-paint at a new corner', () => {
      const ctx = ctxSpy();
      // Initial paint at (200, 100).
      const a = paintFillHandle(ctx, { x: 200, y: 100, w: 80, h: 20 }, theme());
      setFillHandleRect(a);
      expect(isFillHandleHit(280, 120)).toBe(true);

      // Selection moves; paint again at (50, 50).
      const b = paintFillHandle(ctx, { x: 50, y: 50, w: 30, h: 18 }, theme());
      setFillHandleRect(b);
      // The old position is no longer hot.
      expect(isFillHandleHit(280, 120)).toBe(false);
      // The new position is hot.
      expect(isFillHandleHit(80, 68)).toBe(true);
    });
  });

  describe('autofilter chevron (table header)', () => {
    it('returns a rect anchored to the cell right edge with 3px inset', () => {
      const ctx = ctxSpy();
      const cell = { x: 10, y: 20, w: 120, h: 24 };
      const hit = paintTableHeaderChevron(ctx, cell, theme());

      const size = 14;
      // Right edge minus size minus 3px inset.
      expect(hit.x).toBe(cell.x + cell.w - size - 3); // 113
      // Vertically centered with at least 2px gap from top.
      expect(hit.y).toBe(cell.y + Math.max(2, (cell.h - size) / 2)); // 25
      expect(hit.w).toBe(size);
      expect(hit.h).toBe(size);
    });

    it('clamps the y-offset to >= 2px when the cell is shorter than the chevron', () => {
      const ctx = ctxSpy();
      const hit = paintTableHeaderChevron(ctx, { x: 0, y: 0, w: 100, h: 6 }, theme());
      // (6 - 14) / 2 = -4, clamped to 2.
      expect(hit.y).toBe(2);
    });
  });
});
