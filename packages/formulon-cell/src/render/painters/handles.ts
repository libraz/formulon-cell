import type { Range } from '../../engine/types.js';
import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';
import { FALLBACK_ACCENT } from './markers.js';

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
  ctx.strokeStyle = theme.accent || FALLBACK_ACCENT;
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
  ctx.fillStyle = theme.accent || FALLBACK_ACCENT;
  ctx.fillRect(x, y, hs, hs);
  return { x: x - 1, y: y - 1, w: hs + 2, h: hs + 2 };
}

/** Side length and cell inset of the table resize grip, in CSS pixels. */
const TABLE_GRIP_SIZE = 4;
const TABLE_GRIP_INSET = 1;
const TABLE_GRIP_THICKNESS = 2;
const TABLE_GRIP_COLOR = '#485cb1';

/** The square the table resize grip occupies inside a table's last cell. */
export function tableGripRect(cell: Rect, rtl = false): Rect {
  return {
    x: rtl ? cell.x + TABLE_GRIP_INSET : cell.x + cell.w - TABLE_GRIP_INSET - TABLE_GRIP_SIZE,
    y: cell.y + cell.h - TABLE_GRIP_INSET - TABLE_GRIP_SIZE,
    w: TABLE_GRIP_SIZE,
    h: TABLE_GRIP_SIZE,
  };
}

/** L-shaped table resize grip in the table's last cell: a vertical bar on the
 *  trailing side and a horizontal bar along the bottom. */
export function paintTableGrip(ctx: CanvasRenderingContext2D, cell: Rect, rtl = false): void {
  const g = tableGripRect(cell, rtl);
  const bar = TABLE_GRIP_THICKNESS;
  ctx.save();
  ctx.fillStyle = TABLE_GRIP_COLOR;
  ctx.fillRect(rtl ? g.x : g.x + g.w - bar, g.y, bar, g.h);
  ctx.fillRect(g.x, g.y + g.h - bar, g.w, bar);
  ctx.restore();
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

/** Whether a fill preview is the kept part of `selection`: a fill pulled back
 *  inside its source rather than an extension. */
export function isFillShrinkPreview(selection: Range, preview: Range): boolean {
  return (
    preview.sheet === selection.sheet &&
    preview.r0 === selection.r0 &&
    preview.c0 === selection.c0 &&
    preview.r1 <= selection.r1 &&
    preview.c1 <= selection.c1 &&
    (preview.r1 < selection.r1 || preview.c1 < selection.c1)
  );
}

/** White veil over the cells a shrinking fill drops: the selection tint seen
 *  through it reads about a quarter lighter, as measured on the desktop app. */
export function paintFillShrinkShade(ctx: CanvasRenderingContext2D, bounds: Rect): void {
  ctx.save();
  ctx.fillStyle = 'rgba(255, 255, 255, 0.245)';
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  ctx.restore();
}

/** Excel Mac's drag target band, in points: 3 wide centred on the range's
 *  gridlines, light (192,213,199) round a 1 pt (162,183,169) centre line.
 *  Fill-handle and drag-move previews share it. */
const MOVE_BAND_LIGHT = '#c0d5c7';
const MOVE_BAND_CENTRE = '#a2b7a9';

export function paintMoveBand(ctx: CanvasRenderingContext2D, bounds: Rect, pt: number): void {
  const { x, y, w, h } = bounds;
  ctx.save();
  ctx.fillStyle = MOVE_BAND_LIGHT;
  ctx.fillRect(x - 2 * pt, y - 2 * pt, 3 * pt, h + 3 * pt);
  ctx.fillRect(x + w - 2 * pt, y - 2 * pt, 3 * pt, h + 3 * pt);
  ctx.fillRect(x - 2 * pt, y - 2 * pt, w + 3 * pt, 3 * pt);
  ctx.fillRect(x - 2 * pt, y + h - 2 * pt, w + 3 * pt, 3 * pt);
  ctx.fillStyle = MOVE_BAND_CENTRE;
  ctx.fillRect(x - pt, y - pt, pt, h + pt);
  ctx.fillRect(x + w - pt, y - pt, pt, h + pt);
  ctx.fillRect(x - pt, y - pt, w + pt, pt);
  ctx.fillRect(x - pt, y + h - pt, w + pt, pt);
  ctx.restore();
}

/** The Mac marquee: a 2 pt green dash band just inside the cell's outer 1 pt
 *  line, 7 pt dash and 2 pt white gap, the gaps travelling 12 pt/s. */
const MAC_MARQUEE_GREEN = '#217346';
const MAC_MARQUEE_DASH_PT = [7, 2] as const;
const MAC_MARQUEE_PERIOD_PT = 9;
const MAC_MARQUEE_PT_PER_SECOND = 12;

/** The marquee's dash phase at `nowMs`: pixels of travel on the default look,
 *  points of travel on the Mac one. */
export function copyMarqueePhase(nowMs: number, mac: boolean): number {
  if (mac) return ((nowMs / 1000) * MAC_MARQUEE_PT_PER_SECOND) % MAC_MARQUEE_PERIOD_PT;
  return (nowMs / 180) % 7;
}

/** Animated dashed copy-source marquee ("marching ants"). `macPtPx` (the Mac
 *  pt-to-px scale) selects the Mac look, whose `phase` is in points. */
export function paintCopyMarquee(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  phase = 0,
  macPtPx?: number,
): void {
  if (macPtPx !== undefined) {
    const inset = 2 * macPtPx;
    const w = Math.max(0, bounds.w - 2 * inset);
    const h = Math.max(0, bounds.h - 2 * inset);
    ctx.save();
    ctx.lineWidth = 2 * macPtPx;
    ctx.setLineDash([]);
    ctx.strokeStyle = '#ffffff';
    ctx.strokeRect(bounds.x + inset, bounds.y + inset, w, h);
    ctx.setLineDash(MAC_MARQUEE_DASH_PT.map((n) => n * macPtPx));
    ctx.lineDashOffset = -phase * macPtPx;
    ctx.strokeStyle = MAC_MARQUEE_GREEN;
    ctx.strokeRect(bounds.x + inset, bounds.y + inset, w, h);
    ctx.restore();
    return;
  }
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

/** The warning smart tag beside a blocked spill's anchor: a rounded button
 *  holding a yellow triangle, set before the cell's leading edge from its top
 *  (Excel Mac: 24 x 22 px, 13 px clear of the cell). */
export function paintSpillWarningTag(ctx: CanvasRenderingContext2D, cell: Rect, rtl = false): void {
  const w = 24;
  const h = 22;
  const gap = 13;
  const x = rtl ? cell.x + cell.w + gap : cell.x - gap - w;
  const y = cell.y;
  ctx.save();
  const face = ctx.createLinearGradient(0, y, 0, y + h);
  face.addColorStop(0, '#fbfbfb');
  face.addColorStop(1, '#f5f5f5');
  ctx.fillStyle = face;
  ctx.strokeStyle = '#adadad';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.roundRect(x + 0.5, y + 0.5, w - 1, h - 1, 4);
  ctx.fill();
  ctx.stroke();
  // Triangle 16 x 14, inset 4 px; mark stem 4.5 px over a 2 px dot.
  const tx = x + 4;
  const ty = y + 4;
  ctx.fillStyle = '#f7db6f';
  ctx.strokeStyle = '#c8ad4c';
  ctx.beginPath();
  ctx.moveTo(tx + 8, ty + 0.5);
  ctx.lineTo(tx + 15.5, ty + 13.5);
  ctx.lineTo(tx + 0.5, ty + 13.5);
  ctx.closePath();
  ctx.fill();
  ctx.stroke();
  ctx.fillStyle = '#785a15';
  ctx.fillRect(tx + 7, ty + 4, 2, 4.5);
  ctx.fillRect(tx + 7, ty + 10, 2, 2);
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
