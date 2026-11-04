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

/** The Mac spill accent, measured sRGB (49,102,194) on a formula spill and
 *  (48,101,192) on a blocked PivotTable: one colour. */
const SPILL_ACCENT = '#3065c0';

/** The Mac spill boxes: 1 pt over the range's outer gridlines, a 14 / 3 pt dash
 *  for a blocked spill, and a soft shadow fading ~5 pt outside the edges
 *  (sheet grays 238 to 253). */
const MAC_SPILL_DASH_PT = [14, 3] as const;
const MAC_SPILL_SHADOW = 'rgba(0, 0, 0, 0.07)';
const MAC_SPILL_SHADOW_PT = 5;

function strokeMacSpillBox(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  ptPx: number,
  dashed: boolean,
): void {
  ctx.save();
  ctx.strokeStyle = SPILL_ACCENT;
  ctx.lineWidth = ptPx;
  ctx.setLineDash(dashed ? MAC_SPILL_DASH_PT.map((pt) => pt * ptPx) : []);
  ctx.shadowColor = MAC_SPILL_SHADOW;
  ctx.shadowBlur = MAC_SPILL_SHADOW_PT * ptPx;
  // The leading gridlines sit just before the range; the trailing ones are its own.
  ctx.strokeRect(bounds.x - 0.5, bounds.y - 0.5, bounds.w, bounds.h);
  ctx.restore();
}

/** Outline a dynamic-array spill range so it reads as one result. With
 *  `macPtPx` (the Mac pt-to-px scale) it takes the measured Mac look; otherwise
 *  a faint 1px accent ring. */
export function paintSpillOutline(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  macPtPx?: number,
): void {
  if (macPtPx !== undefined) {
    strokeMacSpillBox(ctx, bounds, macPtPx, false);
    return;
  }
  ctx.save();
  ctx.strokeStyle = theme.accent;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  ctx.globalAlpha = 0.65;
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Outline the area a blocked spill (a formula's or a PivotTable's) needs,
 *  with the dashed blue box. `macPtPx` gives the measured Mac geometry. */
export function paintSpillTarget(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  macPtPx?: number,
): void {
  if (macPtPx !== undefined) {
    strokeMacSpillBox(ctx, bounds, macPtPx, true);
    return;
  }
  ctx.save();
  ctx.strokeStyle = SPILL_ACCENT;
  ctx.lineWidth = 1;
  ctx.setLineDash([4, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** The warning smart tag beside a blocked PivotTable's anchor: a yellow
 *  triangle with an exclamation mark, just past the cell's trailing edge. */
export function paintSpillWarningTag(ctx: CanvasRenderingContext2D, cell: Rect): void {
  const size = 12;
  const x = cell.x + cell.w + 2;
  const y = cell.y + (cell.h - size) / 2;
  ctx.save();
  ctx.fillStyle = '#f5c242';
  ctx.strokeStyle = '#a77d10';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(x + size / 2, y + 0.5);
  ctx.lineTo(x + size - 0.5, y + size - 0.5);
  ctx.lineTo(x + 0.5, y + size - 0.5);
  ctx.closePath();
  ctx.fill();
  ctx.stroke();
  ctx.fillStyle = '#1f1f1f';
  ctx.fillRect(x + size / 2 - 0.75, y + 4, 1.5, 4);
  ctx.fillRect(x + size / 2 - 0.75, y + 9, 1.5, 1.5);
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
