import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';

/** Active cell outline. Drawn in a separate pass after all cell text so the
 *  outline never gets clipped by neighbouring cell rects.
 *
 *  Desktop spreadsheets paint the active cell with a crisp ~2px green border and no
 *  inner fill or outer halo. */
export function paintActiveCellOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 2;
  ctx.setLineDash([]);
  ctx.strokeRect(
    Math.round(bounds.x) + 1,
    Math.round(bounds.y) + 1,
    Math.max(0, Math.round(bounds.w) - 2),
    Math.max(0, Math.round(bounds.h) - 2),
  );
  ctx.restore();
}

/** Visible side length of the fill handle in CSS pixels. Desktop spreadsheets use a
 *  small accent-coloured square at the bottom-right of the active selection
 *  range; the user grabs it to drag-fill into adjacent cells. */
export const FILL_HANDLE_SIZE = 6;

/** Spreadsheet-style fill handle. Drawn at the selection range's bottom-right
 *  corner (or active-cell corner when selection is a single cell). The square
 *  is filled in `theme.accent` (resolves from `--fc-accent`, falls back to
 *  `#217346` when unset). */
export function paintFillHandle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const hs = FILL_HANDLE_SIZE;
  // Centre the visible square on the cell's bottom-right corner so half the
  // handle bleeds outside the selection — matches the spreadsheet convention.
  const x = bounds.x + bounds.w - hs / 2;
  const y = bounds.y + bounds.h - hs / 2;
  const accent = theme.accent || '#0078d4';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(x - 1, y - 1, hs + 2, hs + 2);
  ctx.fillStyle = accent;
  ctx.fillRect(x, y, hs, hs);
  return { x: x - 1, y: y - 1, w: hs + 2, h: hs + 2 };
}

/** "Marching ants" outline for the fill drag preview. We omit the actual
 *  ant animation to keep the renderer stateless; a dashed strong rule is
 *  enough to read as a destination. */
export function paintFillPreview(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 1.5;
  ctx.setLineDash([4, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Animated dashed copy-source marquee ("marching ants"). */
export function paintCopyMarquee(ctx: CanvasRenderingContext2D, bounds: Rect, phase = 0): void {
  ctx.save();
  ctx.lineWidth = 1;
  ctx.setLineDash([4, 3]);
  ctx.lineDashOffset = -phase;
  ctx.strokeStyle = '#ffffff';
  ctx.strokeRect(
    bounds.x + 1.5,
    bounds.y + 1.5,
    Math.max(0, bounds.w - 3),
    Math.max(0, bounds.h - 3),
  );
  ctx.lineDashOffset = 3.5 - phase;
  ctx.strokeStyle = '#111111';
  ctx.strokeRect(
    bounds.x + 0.5,
    bounds.y + 0.5,
    Math.max(0, bounds.w - 1),
    Math.max(0, bounds.h - 1),
  );
  ctx.restore();
}

/** Outline a dynamic-array spill range — spreadsheets paint a 1px accent ring
 *  around the spilled rectangle so it reads as a single result. */
export function paintSpillOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  ctx.globalAlpha = 0.65;
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Highlight a cell that is blocking a #SPILL! result with a red dashed
 *  outline. Mirrors the "spill obstruction" indicator. */
export function paintSpillBlocker(ctx: CanvasRenderingContext2D, bounds: Rect): void {
  ctx.save();
  ctx.strokeStyle = '#d83b3b';
  ctx.lineWidth = 1.5;
  ctx.setLineDash([3, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}
