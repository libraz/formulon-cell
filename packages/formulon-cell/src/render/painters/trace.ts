import { MAC_REF_HIGHLIGHT_COLORS, REF_HIGHLIGHT_COLORS } from '../../commands/refs.js';
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
 *  is being edited. Used for "trace precedents while typing" affordance.
 *  With `macPt` (CSS px per point) it draws the Mac desktop's look instead. */
export function paintRefHighlight(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  colorIndex: number,
  macPt?: number,
): void {
  if (macPt !== undefined) {
    paintMacRefHighlight(ctx, bounds, colorIndex, macPt);
    return;
  }
  const color = REF_HIGHLIGHT_COLORS[colorIndex % REF_HIGHLIGHT_COLORS.length] ?? '#1f7ae0';
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 2;
  ctx.setLineDash([5, 3]);
  ctx.strokeRect(bounds.x + 0.5, bounds.y + 0.5, bounds.w - 1, bounds.h - 1);
  ctx.restore();
}

/** Excel Mac: a 2 pt frame over the range's gridlines (the leading edges
 *  just outside, the trailing ones centred on them), a 10% tint 1 pt inside
 *  it and 5 pt corner handles centred 0.5 pt before each corner. */
function paintMacRefHighlight(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  colorIndex: number,
  pt: number,
): void {
  const color = MAC_REF_HIGHLIGHT_COLORS[colorIndex % MAC_REF_HIGHLIGHT_COLORS.length] ?? '#306ac8';
  const { x, y, w, h } = bounds;
  const red = Number.parseInt(color.slice(1, 3), 16);
  const green = Number.parseInt(color.slice(3, 5), 16);
  const blue = Number.parseInt(color.slice(5, 7), 16);
  ctx.save();
  ctx.fillStyle = `rgba(${red}, ${green}, ${blue}, 0.1)`;
  ctx.fillRect(x + pt, y + pt, Math.max(0, w - 3 * pt), Math.max(0, h - 3 * pt));
  ctx.fillStyle = color;
  const outerW = w + 3 * pt;
  const outerH = h + 3 * pt;
  ctx.fillRect(x - 2 * pt, y - 2 * pt, 2 * pt, outerH);
  ctx.fillRect(x + w - pt, y - 2 * pt, 2 * pt, outerH);
  ctx.fillRect(x - 2 * pt, y - 2 * pt, outerW, 2 * pt);
  ctx.fillRect(x - 2 * pt, y + h - pt, outerW, 2 * pt);
  const handle = 5 * pt;
  for (const cx of [x - pt / 2, x + w - pt / 2]) {
    for (const cy of [y - pt / 2, y + h - pt / 2]) {
      ctx.fillRect(cx - handle / 2, cy - handle / 2, handle, handle);
    }
  }
  ctx.restore();
}

/** Excel's colours for the parts of a selected chart's data range. */
const CHART_SOURCE_COLORS = {
  name: '#e0301e',
  categories: '#7030a0',
  values: '#0070c0',
} as const;

/** Outline one part of the selected chart's data range, with the corner
 *  handles Excel draws on it. */
export function paintChartSourceHighlight(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  role: keyof typeof CHART_SOURCE_COLORS,
): void {
  const color = CHART_SOURCE_COLORS[role];
  const handle = 5;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
  ctx.lineWidth = 1.5;
  ctx.strokeRect(bounds.x + 1, bounds.y + 1, bounds.w - 2, bounds.h - 2);
  for (const [x, y] of [
    [bounds.x, bounds.y],
    [bounds.x + bounds.w, bounds.y],
    [bounds.x, bounds.y + bounds.h],
    [bounds.x + bounds.w, bounds.y + bounds.h],
  ] as const) {
    ctx.fillRect(x - handle / 2, y - handle / 2, handle, handle);
  }
  ctx.restore();
}
