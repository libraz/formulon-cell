import { REF_HIGHLIGHT_COLORS } from '../../commands/refs.js';
import type { Rect } from '../geometry.js';

/** Trace-arrow accent colors. Spreadsheets paint precedents in blue and dependents
 *  in a slightly redder hue so the two relations are visually distinct even
 *  when both are active simultaneously. */
export const TRACE_PRECEDENT_COLOR = '#1f7ae0';
export const TRACE_DEPENDENT_COLOR = '#cf3a4c';

/** Paint a small filled dot at the source cell of a trace arrow. The dot
 *  sits at the center of `rect`. Mirrors the blue/red round endpoint convention. */
export function paintTraceDot(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  color: string,
  radius = 3,
): void {
  ctx.save();
  ctx.fillStyle = color;
  ctx.beginPath();
  ctx.arc(rect.x + rect.w / 2, rect.y + rect.h / 2, radius, 0, Math.PI * 2);
  ctx.fill();
  ctx.restore();
}

/** Paint a 1px line from `fromRect` center to `toRect` center, capped with
 *  a small filled triangle at the destination. The arrow head points at
 *  `toRect`'s center so the user reads the direction of dependency. */
export function paintTraceArrow(
  ctx: CanvasRenderingContext2D,
  fromRect: Rect,
  toRect: Rect,
  color: string,
): void {
  const sx = fromRect.x + fromRect.w / 2;
  const sy = fromRect.y + fromRect.h / 2;
  const tx = toRect.x + toRect.w / 2;
  const ty = toRect.y + toRect.h / 2;
  const dx = tx - sx;
  const dy = ty - sy;
  const len = Math.hypot(dx, dy);
  if (len < 0.5) return;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  // Pull back slightly so the line tip butts up against the arrow head's
  // base rather than poking through it.
  const headLen = 9;
  const headHalfWidth = 4;
  const ux = dx / len;
  const uy = dy / len;
  const baseX = tx - ux * headLen;
  const baseY = ty - uy * headLen;
  ctx.beginPath();
  ctx.moveTo(sx, sy);
  ctx.lineTo(baseX, baseY);
  ctx.stroke();
  // Arrow head — filled triangle perpendicular to the line direction.
  const px = -uy;
  const py = ux;
  ctx.beginPath();
  ctx.moveTo(tx, ty);
  ctx.lineTo(baseX + px * headHalfWidth, baseY + py * headHalfWidth);
  ctx.lineTo(baseX - px * headHalfWidth, baseY - py * headHalfWidth);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/** Paint a colored dashed border around a referenced range while a formula
 *  is being edited. Used for "trace precedents while typing" affordance. */
export function paintRefHighlight(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  colorIndex: number,
): void {
  const color = REF_HIGHLIGHT_COLORS[colorIndex % REF_HIGHLIGHT_COLORS.length] ?? '#1f7ae0';
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 2;
  ctx.setLineDash([5, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}
