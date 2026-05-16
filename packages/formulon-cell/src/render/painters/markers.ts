import type { ResolvedTheme } from '../../theme/resolve.js';
import type { Rect } from '../geometry.js';

/** Paint a small ▼ chevron at the right edge of the cell to indicate the
 *  cell has a list validation. Returned rect is the click hit-area. */
export function paintValidationChevron(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const w = 18;
  const h = Math.min(bounds.h, 22);
  const x = bounds.x + bounds.w - w;
  const y = bounds.y + (bounds.h - h) / 2;
  ctx.save();
  ctx.fillStyle = theme.bgRail;
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = theme.rule;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  ctx.fillStyle = theme.fg;
  ctx.beginPath();
  const cx = x + w / 2;
  const cy = y + h / 2;
  ctx.moveTo(cx - 4, cy - 2);
  ctx.lineTo(cx + 4, cy - 2);
  ctx.lineTo(cx, cy + 3);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
  return { x, y, w, h };
}

/** Paint the spreadsheet Table header filter/dropdown affordance. This is visual
 *  only today; table filtering still routes through the normal filter model. */
export function paintTableHeaderChevron(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  theme: ResolvedTheme,
): Rect {
  const size = 14;
  const x = bounds.x + bounds.w - size - 3;
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
): Rect {
  const size = ERROR_TRIANGLE_SIZE;
  const x = bounds.x;
  const y = bounds.y;
  ctx.save();
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.moveTo(x, y);
  ctx.lineTo(x + size, y);
  ctx.lineTo(x, y + size);
  ctx.closePath();
  ctx.fill();
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
): Rect {
  return paintErrorTriangle(ctx, bounds, color);
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
): void {
  const w = 8;
  const h = 9;
  const x = Math.round(bounds.x + bounds.w - w - 2);
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

/** Paint a small filled triangle in the upper-right of the cell to indicate
 *  an attached comment (spreadsheet convention). */
export function paintCommentMarker(ctx: CanvasRenderingContext2D, bounds: Rect): void {
  const size = 5;
  const x = bounds.x + bounds.w - size;
  const y = bounds.y;
  ctx.save();
  ctx.fillStyle = '#d24545';
  ctx.beginPath();
  ctx.moveTo(x + size, y);
  ctx.lineTo(x + size, y + size);
  ctx.lineTo(x, y);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}
