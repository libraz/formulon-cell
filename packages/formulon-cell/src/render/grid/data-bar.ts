import { argbFromCssColor } from '../../engine/theme-color.js';
import type { ConditionalCellOverlay } from '../conditional.js';
import type { Rect } from '../geometry.js';

/** Share of white at the far end of a gradient bar, measured against the
 *  spreadsheet (the bar never fades to pure white). */
const GRADIENT_END_WHITE = 0.82;

/** `color` mixed toward white; null for a colour that is not `#hex` or `rgb()`. */
const towardWhite = (color: string, share: number): string | null => {
  const argb = argbFromCssColor(color);
  const rgb =
    argb !== null
      ? [(argb >>> 16) & 0xff, (argb >>> 8) & 0xff, argb & 0xff]
      : color
          .match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i)
          ?.slice(1, 4)
          .map(Number);
  if (rgb?.length !== 3) return null;
  const [r, g, b] = rgb.map((v) => Math.round(v + (255 - v) * share));
  return `rgb(${r}, ${g}, ${b})`;
};

export function paintDataBar(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  overlay: ConditionalCellOverlay,
): void {
  if (overlay.bar === undefined || !overlay.barColor) return;
  // Measured at 100%: the bar track starts 1px inside the cell, stops 2px
  // short of its right edge, and leaves 1px above and 3px below.
  const left = bounds.x + 1;
  const track = Math.max(0, bounds.w - 3);
  const top = bounds.y + 1;
  const height = Math.max(0, bounds.h - 4);
  const w = track * overlay.bar;
  const axisX = left + track * (overlay.barAxis ?? 0);
  const growsLeft = overlay.barDirection === 'left';
  const endX = growsLeft ? axisX - w : axisX + w;
  ctx.save();
  if (overlay.barGradient) {
    const gradient = ctx.createLinearGradient(axisX, 0, endX, 0);
    gradient.addColorStop(0, overlay.barColor);
    gradient.addColorStop(
      1,
      towardWhite(overlay.barColor, GRADIENT_END_WHITE) ?? 'rgba(255,255,255,0.25)',
    );
    ctx.fillStyle = gradient;
  } else {
    ctx.fillStyle = overlay.barColor;
  }
  ctx.globalAlpha = 1;
  const x = growsLeft ? endX : axisX;
  ctx.fillRect(x, top, w, height);
  if (overlay.barBorderColor && w > 0) {
    ctx.globalAlpha = 1;
    ctx.strokeStyle = overlay.barBorderColor;
    ctx.lineWidth = 1;
    ctx.strokeRect(x, top, w, height);
  }
  if (overlay.barAxisVisible) {
    ctx.globalAlpha = 1;
    ctx.fillStyle = overlay.barAxisColor ?? '#000000';
    ctx.fillRect(axisX - 0.5, top, 1, height);
  }
  ctx.restore();
}
