import type { Rect } from '../geometry.js';

const sameRect = (a: Rect, b: Rect): boolean =>
  a.x === b.x && a.y === b.y && a.w === b.w && a.h === b.h;

export const sameRectList = (a: readonly Rect[], b: readonly Rect[]): boolean =>
  a.length === b.length &&
  a.every((rect, index) => {
    const other = b[index];
    return other !== undefined && sameRect(rect, other);
  });

const rectsTouch = (a: Rect, b: Rect): boolean => {
  const eps = 1e-6;
  const overlapY = a.y < b.y + b.h - eps && b.y < a.y + a.h - eps;
  const overlapX = a.x < b.x + b.w - eps && b.x < a.x + a.w - eps;
  const sideBySide =
    (Math.abs(a.x + a.w - b.x) < eps || Math.abs(b.x + b.w - a.x) < eps) && overlapY;
  const stacked = (Math.abs(a.y + a.h - b.y) < eps || Math.abs(b.y + b.h - a.y) < eps) && overlapX;
  return (overlapX && overlapY) || sideBySide || stacked;
};

/** Join visible range rects that share an edge. Freeze quadrants are separate
 *  for fills and outlines, but their screen coordinates are contiguous, so a
 *  single text clip must span their union. Page-layout gutters remain
 *  separate components and therefore do not make text cross a paper gap. */
export const connectedRects = (rects: readonly Rect[]): Rect[] => {
  const pending = rects.map((rect) => ({ ...rect }));
  let changed = true;
  while (changed) {
    changed = false;
    outer: for (let i = 0; i < pending.length; i += 1) {
      const first = pending[i];
      if (!first) continue;
      for (let j = i + 1; j < pending.length; j += 1) {
        const second = pending[j];
        if (!second || !rectsTouch(first, second)) continue;
        pending[i] = {
          x: Math.min(first.x, second.x),
          y: Math.min(first.y, second.y),
          w: Math.max(first.x + first.w, second.x + second.w) - Math.min(first.x, second.x),
          h: Math.max(first.y + first.h, second.y + second.h) - Math.min(first.y, second.y),
        };
        pending.splice(j, 1);
        changed = true;
        break outer;
      }
    }
  }
  return pending;
};

/** Pick the visible rectangle at a range's logical bottom/trailing corner.
 *  `rangeRects` returns freeze quadrants in logical row/column order, but the
 *  physical trailing edge moves to the left on a right-to-left sheet. */
export const trailingRect = (rects: readonly Rect[], rtl: boolean): Rect | null => {
  if (rects.length === 0) return null;
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.h));
  const bottomRects = rects.filter((rect) => rect.y + rect.h === bottom);
  return bottomRects.reduce(
    (best, rect) => {
      if (!best) return rect;
      if (rtl) return rect.x < best.x ? rect : best;
      return rect.x + rect.w > best.x + best.w ? rect : best;
    },
    null as Rect | null,
  );
};
