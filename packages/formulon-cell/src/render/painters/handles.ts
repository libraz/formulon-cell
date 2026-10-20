import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';

/** Selection outline shared by the active cell and a multi-cell range.
 *
 *  Desktop spreadsheets paint a crisp 2px accent border that covers the
 *  range's own gridlines plus one pixel outside them, so no gridline shows
 *  between the border and the neighbouring cells. Drawn in a separate pass
 *  after all cell text so neighbouring cell rects never clip it. */
export function paintSelectionOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): void {
  const left = Math.round(bounds.x);
  const top = Math.round(bounds.y);
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 2;
  ctx.setLineDash([]);
  ctx.strokeRect(
    left,
    top,
    Math.round(bounds.x + bounds.w) + 1 - left,
    Math.round(bounds.y + bounds.h) + 1 - top,
  );
  ctx.restore();
}

/** Visible side length of the fill handle in CSS pixels: a 5px accent square
 *  centred on the selection's trailing bottom gridline intersection. */
export const FILL_HANDLE_SIZE = 5;

/** Spreadsheet-style fill handle at the selection range's trailing bottom
 *  corner. A 1px white gap separates it from the outline on the two sides
 *  that face into the selection. Returns the square padded by 1px on every
 *  side, so the grab zone stays centred on the corner. */
export function paintFillHandle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  rtl = false,
): Rect {
  const hs = FILL_HANDLE_SIZE;
  const half = Math.floor(hs / 2);
  const cornerX = Math.round(rtl ? bounds.x : bounds.x + bounds.w);
  const x = cornerX - half;
  const y = Math.round(bounds.y + bounds.h) - half;
  const gapX = rtl ? x : x - 1;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(gapX, y - 1, hs + 1, hs + 1);
  ctx.fillStyle = theme.accent || '#0078d4';
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
