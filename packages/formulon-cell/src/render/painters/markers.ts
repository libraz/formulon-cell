import type { ResolvedTheme } from '../../theme/resolve.js';
import { mirrorInRect, type Rect } from '../geometry.js';

/** In-cell dropdown button, measured at 100%: 16x17px, hung from the cell's
 *  trailing edge 1px below its top (covering the outer half of the selection
 *  outline), with a gray frame (no leading side), a vertical face gradient and a 9x6 gray arrow. */
const DROPDOWN_W = 16;
const DROPDOWN_H = 17;
const DROPDOWN_FRAME = '#999999';
/** The selection outline reads darker along the button. */
const DROPDOWN_OUTLINE_SHADE = '#134529';
const DROPDOWN_ARROW = '#838383';
const DROPDOWN_FACE_TOP = '#ffffff';
const DROPDOWN_FACE_BOTTOM = '#f7f7f7';
const DROPDOWN_HIGHLIGHT = '#f0f0f0';

/** Excel Mac's dropdown button, in points: 14 x 14 on the cell's trailing
 *  gridline from 1 above its top, a 255 -> 245 face framed on three sides
 *  and a centred 5 x 4.5 arrow. */
const MAC_DROPDOWN = {
  size: 14,
  face: ['#ffffff', '#f5f5f5'],
  frame: '#999999',
  arrow: '#7d7d7d',
  arrowW: 5,
  arrowH: 4.5,
} as const;

function paintMacDropdownButton(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  rtl: boolean,
  pt: number,
): Rect {
  const size = MAC_DROPDOWN.size * pt;
  const x = rtl ? bounds.x + pt - size : bounds.x + bounds.w - pt;
  const y = bounds.y - pt;
  ctx.save();
  const face = ctx.createLinearGradient(0, y, 0, y + size);
  face.addColorStop(0, MAC_DROPDOWN.face[0]);
  face.addColorStop(1, MAC_DROPDOWN.face[1]);
  ctx.fillStyle = face;
  ctx.fillRect(x, y, size, size);
  // Framed top, outer side and bottom; the cell's own gridline closes it.
  ctx.fillStyle = MAC_DROPDOWN.frame;
  ctx.fillRect(x, y, size, pt / 2);
  ctx.fillRect(rtl ? x : x + size - pt / 2, y, pt / 2, size);
  ctx.fillRect(x, y + size - pt / 2, size, pt / 2);
  const cx = x + size / 2;
  const top = y + (size - MAC_DROPDOWN.arrowH * pt) / 2;
  ctx.fillStyle = MAC_DROPDOWN.arrow;
  ctx.beginPath();
  ctx.moveTo(cx - (MAC_DROPDOWN.arrowW * pt) / 2, top);
  ctx.lineTo(cx + (MAC_DROPDOWN.arrowW * pt) / 2, top);
  ctx.lineTo(cx, top + MAC_DROPDOWN.arrowH * pt);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w: size, h: size };
}

/** Paint the list-validation / totals dropdown button beside the active cell.
 *  Returned rect is the click hit-area. */
export function paintValidationChevron(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  _theme: ResolvedTheme,
  rtl = false,
  macPt?: number,
): Rect {
  if (macPt !== undefined) return paintMacDropdownButton(ctx, bounds, rtl, macPt);
  const w = DROPDOWN_W;
  const h = DROPDOWN_H;
  const x = mirrorInRect(bounds, bounds.x + bounds.w, w, rtl);
  const y = bounds.y + 1;
  ctx.save();
  // The face starts one pixel below the top so the selection outline's
  // corner stays visible; the top frame starts one pixel in.
  const body = ctx.createLinearGradient(0, y + 1, 0, y + h);
  body.addColorStop(0, DROPDOWN_FACE_TOP);
  body.addColorStop(1, DROPDOWN_FACE_BOTTOM);
  ctx.fillStyle = body;
  ctx.fillRect(x, y + 1, w, h - 1);
  ctx.fillStyle = DROPDOWN_HIGHLIGHT;
  ctx.fillRect(rtl ? x : x + 1, y + 1, w - 1, 0.5);
  // No leading border: the active-cell outline abuts the button.
  ctx.strokeStyle = DROPDOWN_FRAME;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(rtl ? x + w - 1 : x + 1, y + 0.5);
  ctx.lineTo(rtl ? x + 0.5 : x + w - 0.5, y + 0.5);
  ctx.lineTo(rtl ? x + 0.5 : x + w - 0.5, y + h - 0.5);
  ctx.lineTo(rtl ? x + w : x, y + h - 0.5);
  ctx.stroke();
  if (!rtl) {
    ctx.fillStyle = DROPDOWN_OUTLINE_SHADE;
    ctx.fillRect(x - 1, y, 2, 1);
    ctx.fillRect(x - 1, y + 1, 1, h - 1);
  }
  ctx.fillStyle = DROPDOWN_ARROW;
  ctx.beginPath();
  ctx.moveTo(x + 3, y + 6.25);
  ctx.lineTo(x + 12, y + 6.25);
  ctx.lineTo(x + 7.5, y + 11.9);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w, h };
}

export const TABLE_HEADER_CHEVRON_SIZE = 14;
export const TABLE_HEADER_CHEVRON_INSET = 3;

/** Paint the spreadsheet Table header filter/dropdown affordance. This is visual
 *  only today; table filtering still routes through the normal filter model. */
export function paintTableHeaderChevron(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  rtl = false,
): Rect {
  const size = TABLE_HEADER_CHEVRON_SIZE;
  const x = mirrorInRect(
    bounds,
    bounds.x + bounds.w - size - TABLE_HEADER_CHEVRON_INSET,
    size,
    rtl,
  );
  const y = bounds.y + Math.max(2, (bounds.h - size) / 2);
  ctx.save();
  ctx.fillStyle = 'rgba(255,255,255,0.72)';
  ctx.fillRect(x, y, size, size);
  ctx.strokeStyle = theme.rule;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, size - 1, size - 1);
  ctx.fillStyle = theme.fgMute || theme.fg;
  ctx.beginPath();
  const cx = x + size / 2;
  const cy = y + size / 2 + 1;
  ctx.moveTo(cx - 3.5, cy - 2);
  ctx.lineTo(cx + 3.5, cy - 2);
  ctx.lineTo(cx, cy + 2.5);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w: size, h: size };
}

/** The error-check smart tag, in points: a square button hung left of the cell. */
const SMART_TAG_SIZE_PT = 22;
const SMART_TAG_GAP_PT = 11;
const SMART_TAG_TOP_PT = 1.5;

/** Paint the error-check smart tag to the leading side of the active cell: a
 *  rounded button holding a yellow warning triangle. Returned rect is the click
 *  hit-area. */
export function paintErrorSmartTag(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  ptPx: number,
  rtl = false,
): Rect {
  const size = SMART_TAG_SIZE_PT * ptPx;
  const gap = SMART_TAG_GAP_PT * ptPx;
  const x = mirrorInRect(bounds, bounds.x - gap - size, size, rtl);
  const y = bounds.y + SMART_TAG_TOP_PT * ptPx;
  const radius = 4 * ptPx;
  ctx.save();
  ctx.beginPath();
  ctx.roundRect(x + 0.5, y + 0.5, size - 1, size - 1, radius);
  ctx.fillStyle = '#fbfbfb';
  ctx.fill();
  ctx.strokeStyle = '#b4b4b4';
  ctx.lineWidth = 1;
  ctx.stroke();
  const cx = x + size / 2;
  const cy = y + size / 2 + 0.5 * ptPx;
  const half = 8 * ptPx;
  const tall = 7 * ptPx;
  ctx.beginPath();
  ctx.moveTo(cx, cy - tall);
  ctx.lineTo(cx + half, cy + tall - 1 * ptPx);
  ctx.lineTo(cx - half, cy + tall - 1 * ptPx);
  ctx.closePath();
  ctx.fillStyle = '#ffd861';
  ctx.fill();
  ctx.strokeStyle = '#dab139';
  ctx.lineWidth = 1.2 * ptPx;
  ctx.lineJoin = 'round';
  ctx.stroke();
  ctx.fillStyle = '#4a3b00';
  ctx.fillRect(cx - 0.8 * ptPx, cy - 3 * ptPx, 1.6 * ptPx, 4.5 * ptPx);
  ctx.fillRect(cx - 0.8 * ptPx, cy + 2.4 * ptPx, 1.6 * ptPx, 1.6 * ptPx);
  ctx.restore();
  return { x, y, w: size, h: size };
}

/** Side of the triangle hot-zone painted by `paintErrorTriangle` /
 *  `paintValidationTriangle`. Click hit-tests in `error-menu` use the same
 *  constant so the visual and clickable areas line up. */
export const ERROR_TRIANGLE_SIZE = 6;

/** Paint a small filled triangle in the upper-LEFT of the cell. Used for
 *  formula-error and data-validation indicators. The triangle sits flush
 *  against the top and left cell edges and points down-right. Returned rect
 *  is the click hit-area (slightly padded so the corner is comfortable to
 *  hit). */
export function paintErrorTriangle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  color: string,
  rtl = false,
  size = ERROR_TRIANGLE_SIZE,
): Rect {
  const x = mirrorInRect(bounds, bounds.x, size, rtl);
  const y = bounds.y;
  ctx.save();
  ctx.fillStyle = color;
  // Right angle at the cell's leading corner; the hypotenuse falls inwards as
  // a whole-pixel staircase, unsmoothed as the spreadsheet draws it.
  for (let row = 0; row < size; row += 1) {
    const w = size - row;
    ctx.fillRect(rtl ? x + size - w : x, y + row, w, 1);
  }
  ctx.restore();
  return { x, y, w: size, h: size };
}

/** Convenience alias — spreadsheets paint DV violations the same shape as formula
 *  errors but red. The grid wires both through `paintErrorTriangle`; this
 *  thin wrapper keeps the call sites self-documenting. */
export function paintValidationTriangle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  color = '#d24545',
  rtl = false,
): Rect {
  return paintErrorTriangle(ctx, bounds, color, rtl);
}

/** Excel-style "Circle Invalid Data" marker. It intentionally sits inside the
 *  cell bounds so it remains visible with gridlines, fills, and selection
 *  outlines. */
export function paintValidationCircle(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  color = '#d24545',
): void {
  const insetX = Math.min(5, Math.max(2, bounds.w * 0.08));
  const insetY = Math.min(4, Math.max(2, bounds.h * 0.12));
  const x = bounds.x + insetX;
  const y = bounds.y + insetY;
  const w = Math.max(4, bounds.w - insetX * 2);
  const h = Math.max(4, bounds.h - insetY * 2);
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
  ctx.stroke();
  ctx.restore();
}

/** Paint a small lock-icon overlay in the upper-right corner of a cell.
 *  Used to flag cells that are still writable when the sheet is otherwise
 *  protected (i.e. `format.locked === false` while the sheet is protected).
 *  The shape is a 7×8 rounded body + arched shackle drawn in
 *  `theme.accent` so it reads as an affordance, not an error. */
export function paintLockMarker(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
  rtl = false,
): void {
  const w = 8;
  const h = 9;
  const x = Math.round(mirrorInRect(bounds, bounds.x + bounds.w - w - 2, w, rtl));
  const y = Math.round(bounds.y + 2);
  const color = theme.accent || '#0078d4';
  ctx.save();
  // Shackle (top arch, drawn as a stroked half-circle).
  ctx.strokeStyle = color;
  ctx.lineWidth = 1.25;
  ctx.beginPath();
  ctx.arc(x + w / 2, y + 3, 2.25, Math.PI, 0);
  ctx.stroke();
  // Body.
  ctx.fillStyle = color;
  ctx.fillRect(x + 1, y + 3, w - 2, h - 4);
  ctx.restore();
}

/** Paint the note indicator: a red right triangle, 6 CSS px a side, in the
 *  cell's trailing top corner inside its gridline. It is a staircase of
 *  one-CSS-px steps (12x12 device px in 84 pixels at 2x), not an antialiased
 *  diagonal. */
export function paintCommentMarker(ctx: CanvasRenderingContext2D, bounds: Rect, rtl = false): void {
  const size = 6;
  const x = mirrorInRect(bounds, bounds.x + bounds.w - 1 - size, size, rtl);
  ctx.save();
  ctx.fillStyle = '#ff0000';
  for (let row = 0; row < size; row += 1) {
    const width = size - row;
    ctx.fillRect(rtl ? x : x + size - width, bounds.y + row, width, 1);
  }
  ctx.restore();
}
