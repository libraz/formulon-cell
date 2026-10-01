import type { ConditionalCellOverlay } from '../conditional.js';
import type { Rect } from '../geometry.js';

export function paintDataBar(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  overlay: ConditionalCellOverlay,
): void {
  if (overlay.bar === undefined || !overlay.barColor) return;
  const w = bounds.w * overlay.bar;
  const axisX = bounds.x + bounds.w * (overlay.barAxis ?? 0);
  const growsLeft = overlay.barDirection === 'left';
  const endX = growsLeft ? axisX - w : axisX + w;
  ctx.save();
  if (overlay.barGradient) {
    const gradient = ctx.createLinearGradient(axisX, 0, endX, 0);
    gradient.addColorStop(0, overlay.barColor);
    gradient.addColorStop(1, 'rgba(255,255,255,0.25)');
    ctx.fillStyle = gradient;
  } else {
    ctx.fillStyle = overlay.barColor;
  }
  ctx.globalAlpha = 1;
  const x = growsLeft ? endX : axisX;
  ctx.fillRect(x, bounds.y + 1, w, bounds.h - 2);
  if (overlay.barBorderColor && w > 0) {
    ctx.globalAlpha = 1;
    ctx.strokeStyle = overlay.barBorderColor;
    ctx.lineWidth = 1;
    ctx.strokeRect(x, bounds.y + 1, w, bounds.h - 2);
  }
  if (overlay.barAxisVisible) {
    ctx.globalAlpha = 1;
    ctx.fillStyle = overlay.barAxisColor ?? '#000000';
    ctx.fillRect(axisX - 0.5, bounds.y + 1, 1, bounds.h - 2);
  }
  ctx.restore();
}
