import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';

/** Outer box of the 2px selection outline around `bounds`. Desktop
 *  spreadsheets cover the range's own gridlines plus one pixel outside them,
 *  so no gridline shows between the border and the neighbouring cells. A cell
 *  owns the rule on its trailing edges, so the range's leading rules sit just
 *  before it — the column rules sit after each edge instead in RTL. The cell
 *  editor takes the same box so editing keeps the frame. */
export function selectionOutlineFrame(bounds: Rect, rtl = false): Rect {
  const left = Math.round(bounds.x);
  const top = Math.round(bounds.y);
  const right = Math.round(bounds.x + bounds.w);
  const bottom = Math.round(bounds.y + bounds.h);
  const x = rtl ? left - 1 : left - 2;
  const y = top - 2;
  return { x, y, w: (rtl ? right + 2 : right + 1) - x, h: bottom + 1 - y };
}

/** Selection outline shared by the active cell and a multi-cell range,
 *  stroked inside `selectionOutlineFrame`. A shaded range keeps a 1px
 *  sheet-coloured gap inside the border (`inset`). Drawn in a separate pass
 *  after all cell text so neighbouring cell rects never clip it. */
export function paintSelectionOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  inset = false,
  rtl = false,
  limit?: { x?: number; y?: number },
): void {
  const raw = selectionOutlineFrame(bounds, rtl);
  // With no header strip the frame of a full-axis selection stays inside the grid.
  const x = Math.max(raw.x, limit?.x ?? raw.x);
  const y = Math.max(raw.y, limit?.y ?? raw.y);
  const frame = { x, y, w: raw.x + raw.w - x, h: raw.y + raw.h - y };
  // Centre lines of the 2px border.
  const x0 = frame.x + 1;
  const y0 = frame.y + 1;
  const w = frame.w - 2;
  const h = frame.h - 2;
  ctx.save();
  ctx.setLineDash([]);
  if (inset && w > 4 && h > 4) {
    ctx.strokeStyle = theme.bg;
    ctx.lineWidth = 1;
    ctx.strokeRect(x0 + 1.5, y0 + 1.5, w - 3, h - 3);
  }
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 2;
  ctx.strokeRect(x0, y0, w, h);
  ctx.restore();
}

/** Rule of a range in a multi-range selection (measured light grey). */
const MULTI_RANGE_RULE = '#a5a5a5';

/** Outline of one range of a multi-range selection: a one-pixel grey rule on
 *  the range's gridlines with a one-pixel sheet-coloured inset, in place of
 *  the single-range accent frame. */
export function paintRangeOutlineThin(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  rtl = false,
): void {
  const left = Math.round(bounds.x);
  const top = Math.round(bounds.y);
  const right = Math.round(bounds.x + bounds.w);
  const bottom = Math.round(bounds.y + bounds.h);
  const x = rtl ? left : left - 1;
  const y = top - 1;
  const w = (rtl ? right + 1 : right) - x;
  const h = bottom - y;
  ctx.save();
  ctx.setLineDash([]);
  ctx.lineWidth = 1;
  ctx.strokeStyle = MULTI_RANGE_RULE;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  if (w > 4 && h > 4) {
    ctx.strokeStyle = theme.bg;
    ctx.strokeRect(x + 1.5, y + 1.5, w - 3, h - 3);
  }
  ctx.restore();
}

/** Box around the active cell inside a multi-range selection: one accent
 *  pixel, one pixel inside the cell's surface. */
export function paintActiveCellBox(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  rtl = false,
): void {
  const left = Math.round(bounds.x);
  const top = Math.round(bounds.y);
  const right = Math.round(bounds.x + bounds.w);
  const bottom = Math.round(bounds.y + bounds.h);
  ctx.save();
  ctx.setLineDash([]);
  ctx.lineWidth = 1;
  ctx.strokeStyle = theme.accent || '#0078d4';
  ctx.strokeRect((rtl ? left + 1 : left) + 1.5, top + 1.5, right - left - 4, bottom - top - 4);
  ctx.restore();
}

/** Visible side length of the fill handle in CSS pixels: a 5px accent square
 *  centred on the selection's trailing bottom gridline intersection. */
export const FILL_HANDLE_SIZE = 5;

/** Spreadsheet-style fill handle at the selection range's trailing bottom
 *  corner. A 1px white gap separates it from the outline on the two sides
 *  that face into the selection. Returns the square padded by 1px on every
 *  side, so the grab zone stays centred on the corner. `ring` puts the white
 *  gap on all four sides, for a handle moved off the bottom-right corner. */
export function paintFillHandle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  rtl = false,
  ring = false,
): Rect {
  const hs = FILL_HANDLE_SIZE;
  const half = Math.floor(hs / 2);
  // The trailing corner rule ends at the boundary (starts at it in RTL).
  const cornerX = Math.round(rtl ? bounds.x : bounds.x + bounds.w);
  const x = rtl ? cornerX - half : cornerX - half - 1;
  const y = Math.round(bounds.y + bounds.h) - half - 1;
  const gapX = rtl ? x : x - 1;
  ctx.fillStyle = '#ffffff';
  if (ring) ctx.fillRect(x - 1, y - 1, hs + 2, hs + 2);
  else ctx.fillRect(gapX, y - 1, hs + 1, hs + 1);
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
