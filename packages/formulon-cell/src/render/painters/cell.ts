import type { ConditionalIconSet } from '../../store/store.js';
import type { Rect } from '../geometry.js';
import type { CellPaintCtx } from './types.js';

/* ---- Base painters wired in MS-A. Future slots (formatting, CF, DV,
 * comments, hyperlinks) live next to these, capability-gated. ---- */

export function paintCellBackground({
  ctx,
  bounds,
  theme,
  isInRange,
  isActive,
}: CellPaintCtx): void {
  if (isActive) {
    ctx.fillStyle = theme.bgElev;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
    return;
  }
  if (isInRange) {
    ctx.fillStyle = theme.accentSoft;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  }
  // else: grid bg painted globally by paintGridSurface; do nothing.
}

/** Width of the left-gutter that hosts conditional-format icon-set glyphs.
 *  Cell text is right-shifted by this amount when an icon overlay is
 *  present so the artwork and value don't overlap. */
export const CONDITIONAL_ICON_GUTTER = 16;

/** Paint a small icon-set glyph at the left of a cell rect. The artwork is
 *  drawn with primitive shapes — no SVG / image asset — so the renderer
 *  stays self-contained. `slot` is 0-based; the family decides the count
 *  (3 or 5) and how the slot maps to color/direction. */
export function paintConditionalIcon(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  kind: ConditionalIconSet,
  slot: number,
): void {
  const cx = rect.x + CONDITIONAL_ICON_GUTTER / 2;
  const cy = rect.y + rect.h / 2;
  const r = Math.min(5, Math.floor(rect.h * 0.32));
  ctx.save();
  if (kind === 'arrows3' || kind === 'arrows5') {
    // Map slot to a unit-vector direction. arrows3: [down, right, up];
    // arrows5: [down, down-right, right, up-right, up]. Color tracks slot:
    // low → red, mid → amber, high → green. arrows5 reuses the same ramp
    // with finer steps.
    const slots = kind === 'arrows5' ? 5 : 3;
    const idx = Math.max(0, Math.min(slots - 1, slot));
    const colors =
      kind === 'arrows5'
        ? ['#d24545', '#e07a4d', '#cfa64a', '#7fb352', '#3aa055']
        : ['#d24545', '#cfa64a', '#3aa055'];
    const color = colors[idx] ?? '#777';
    // Direction angle in radians, 0 = right, -π/2 = up, π/2 = down.
    const angles =
      kind === 'arrows5'
        ? [Math.PI / 2, Math.PI / 4, 0, -Math.PI / 4, -Math.PI / 2]
        : [Math.PI / 2, 0, -Math.PI / 2];
    const a = angles[idx] ?? 0;
    const len = r + 2;
    const dx = Math.cos(a) * len;
    const dy = Math.sin(a) * len;
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(cx - dx, cy - dy);
    ctx.lineTo(cx + dx, cy + dy);
    ctx.stroke();
    // Arrow head — small triangle perpendicular to the direction.
    const px = -Math.sin(a);
    const py = Math.cos(a);
    const headLen = 3;
    const headHalf = 2.5;
    const baseX = cx + dx - Math.cos(a) * headLen;
    const baseY = cy + dy - Math.sin(a) * headLen;
    ctx.beginPath();
    ctx.moveTo(cx + dx, cy + dy);
    ctx.lineTo(baseX + px * headHalf, baseY + py * headHalf);
    ctx.lineTo(baseX - px * headHalf, baseY - py * headHalf);
    ctx.closePath();
    ctx.fill();
  } else if (kind === 'triangles3') {
    const idx = Math.max(0, Math.min(2, slot));
    const colors = ['#d24545', '#cfa64a', '#3aa055'];
    ctx.fillStyle = colors[idx] ?? '#777';
    if (idx === 1) {
      ctx.fillRect(cx - r, cy - 1, r * 2, 2);
    } else {
      const up = idx === 2;
      ctx.beginPath();
      ctx.moveTo(cx, cy + (up ? -r : r));
      ctx.lineTo(cx - r, cy + (up ? r : -r));
      ctx.lineTo(cx + r, cy + (up ? r : -r));
      ctx.closePath();
      ctx.fill();
    }
  } else if (kind === 'traffic3' || kind === 'trafficRim3') {
    // Three colored circles, only the slot's circle is filled solid; the
    // other two are dim outlines so the icon reads as a single state.
    const idx = Math.max(0, Math.min(2, slot));
    const colors = ['#d24545', '#cfa64a', '#3aa055'];
    const color = colors[idx] ?? '#777';
    if (kind === 'trafficRim3') {
      ctx.strokeStyle = '#4f4f4f';
      ctx.lineWidth = 1.2;
      ctx.beginPath();
      ctx.rect(cx - r - 2, cy - r - 2, (r + 2) * 2, (r + 2) * 2);
      ctx.stroke();
    }
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fill();
  } else if (kind === 'symbols3') {
    const idx = Math.max(0, Math.min(2, slot));
    ctx.lineWidth = 1.6;
    if (idx === 2) {
      ctx.strokeStyle = '#3aa055';
      ctx.beginPath();
      ctx.moveTo(cx - r, cy);
      ctx.lineTo(cx - 1, cy + r - 1);
      ctx.lineTo(cx + r, cy - r);
      ctx.stroke();
    } else if (idx === 1) {
      ctx.strokeStyle = '#cfa64a';
      ctx.beginPath();
      ctx.moveTo(cx, cy - r);
      ctx.lineTo(cx, cy + 1);
      ctx.stroke();
      ctx.fillStyle = '#cfa64a';
      ctx.beginPath();
      ctx.arc(cx, cy + r - 1, 1.2, 0, Math.PI * 2);
      ctx.fill();
    } else {
      ctx.strokeStyle = '#d24545';
      ctx.beginPath();
      ctx.moveTo(cx - r, cy - r);
      ctx.lineTo(cx + r, cy + r);
      ctx.moveTo(cx + r, cy - r);
      ctx.lineTo(cx - r, cy + r);
      ctx.stroke();
    }
  } else if (kind === 'flags3') {
    const idx = Math.max(0, Math.min(2, slot));
    const colors = ['#d24545', '#cfa64a', '#3aa055'];
    ctx.strokeStyle = '#5c5c5c';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(cx - r, cy + r);
    ctx.lineTo(cx - r, cy - r);
    ctx.stroke();
    ctx.fillStyle = colors[idx] ?? '#777';
    ctx.beginPath();
    ctx.moveTo(cx - r, cy - r);
    ctx.lineTo(cx + r, cy - r + 1.5);
    ctx.lineTo(cx - r, cy + 1);
    ctx.closePath();
    ctx.fill();
  } else if (kind === 'quarters5' || kind === 'ratings5') {
    const idx = Math.max(0, Math.min(4, slot));
    ctx.strokeStyle = '#5c5c5c';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.stroke();
    ctx.fillStyle = '#3f7ad8';
    if (kind === 'ratings5') {
      ctx.beginPath();
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
      ctx.globalAlpha = 0.25 + idx * 0.18;
      ctx.fill();
      ctx.globalAlpha = 1;
    } else {
      const end = -Math.PI / 2 + ((idx + 1) / 4) * Math.PI * 2;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, -Math.PI / 2, end);
      ctx.closePath();
      ctx.fill();
    }
  } else if (kind === 'bars5' || kind === 'boxes5') {
    const idx = Math.max(0, Math.min(4, slot));
    ctx.fillStyle = '#3f7ad8';
    if (kind === 'bars5') {
      const barW = 2;
      const gap = 1;
      const baseX = cx - 5;
      for (let i = 0; i < 5; i += 1) {
        const h = 3 + i * 2;
        ctx.globalAlpha = i <= idx ? 1 : 0.25;
        ctx.fillRect(baseX + i * (barW + gap), cy + r - h, barW, h);
      }
      ctx.globalAlpha = 1;
    } else {
      const size = 3;
      const baseX = cx - 5;
      const baseY = cy - 4;
      for (let i = 0; i < 5; i += 1) {
        ctx.globalAlpha = i <= idx ? 1 : 0.25;
        ctx.fillRect(baseX + (i % 3) * 4, baseY + Math.floor(i / 3) * 4, size, size);
      }
      ctx.globalAlpha = 1;
    }
  } else {
    // stars3 — outlined or filled 5-pointed star. slot 0 = empty,
    // slot 1 = half (filled lower-half), slot 2 = full.
    const idx = Math.max(0, Math.min(2, slot));
    const star = (filled: boolean): void => {
      ctx.beginPath();
      const spikes = 5;
      const outerR = r;
      const innerR = r * 0.45;
      for (let i = 0; i < spikes * 2; i += 1) {
        const angle = (i * Math.PI) / spikes - Math.PI / 2;
        const radius = i % 2 === 0 ? outerR : innerR;
        const px = cx + Math.cos(angle) * radius;
        const py = cy + Math.sin(angle) * radius;
        if (i === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
      }
      ctx.closePath();
      if (filled) ctx.fill();
      else ctx.stroke();
    };
    ctx.strokeStyle = '#cfa64a';
    ctx.fillStyle = '#cfa64a';
    ctx.lineWidth = 1;
    if (idx === 2) {
      star(true);
    } else if (idx === 1) {
      // Half-fill — clip the right half then fill, then outline whole.
      ctx.save();
      ctx.beginPath();
      ctx.rect(rect.x, cy - r - 1, cx - rect.x, (r + 1) * 2);
      ctx.clip();
      star(true);
      ctx.restore();
      star(false);
    } else {
      star(false);
    }
  }
  ctx.restore();
}

/** User-set background fill — drawn above range tint and active highlight so
 *  cells with explicit fill stay visually filled even when selected. */
export function paintCellFill({ ctx, bounds, format }: CellPaintCtx): void {
  const fill = format?.fill;
  const pattern = format?.fillPattern;
  if (!fill && !pattern) return;
  if (fill) {
    ctx.fillStyle = fill;
    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  }
  if (!pattern) return;
  const color = format?.fillPatternColor ?? '#000000';
  ctx.save();
  ctx.beginPath();
  ctx.rect(bounds.x, bounds.y, bounds.w, bounds.h);
  ctx.clip();
  ctx.strokeStyle = color;
  ctx.fillStyle = color;
  ctx.lineWidth = 1;
  switch (pattern) {
    case 'gray125':
    case 'gray25':
    case 'gray50': {
      const step = pattern === 'gray125' ? 6 : pattern === 'gray25' ? 4 : 3;
      const size = pattern === 'gray50' ? 1.5 : 1;
      for (let y = bounds.y + 2; y < bounds.y + bounds.h; y += step) {
        for (let x = bounds.x + 2; x < bounds.x + bounds.w; x += step) {
          ctx.fillRect(x, y, size, size);
        }
      }
      break;
    }
    case 'horizontal':
      for (let y = bounds.y + 3; y < bounds.y + bounds.h; y += 4) {
        ctx.beginPath();
        ctx.moveTo(bounds.x, y);
        ctx.lineTo(bounds.x + bounds.w, y);
        ctx.stroke();
      }
      break;
    case 'vertical':
      for (let x = bounds.x + 3; x < bounds.x + bounds.w; x += 4) {
        ctx.beginPath();
        ctx.moveTo(x, bounds.y);
        ctx.lineTo(x, bounds.y + bounds.h);
        ctx.stroke();
      }
      break;
    case 'diagonalDown':
    case 'diagonalUp': {
      const direction = pattern === 'diagonalDown' ? 1 : -1;
      for (let x = bounds.x - bounds.h; x < bounds.x + bounds.w; x += 6) {
        ctx.beginPath();
        ctx.moveTo(x, direction > 0 ? bounds.y : bounds.y + bounds.h);
        ctx.lineTo(x + bounds.h, direction > 0 ? bounds.y + bounds.h : bounds.y);
        ctx.stroke();
      }
      break;
    }
  }
  ctx.restore();
}

/** Per-cell border lines. Drawn after gridlines + cell text but before the
 *  active outline so user borders sit above the hairline grid. Each side may
 *  be a boolean (legacy thin solid) or an object with `{style, color}`. */
export function paintCellBorders({ ctx, bounds, theme, format }: CellPaintCtx): void {
  const sides = format?.borders;
  if (!sides) return;
  const drawSide = (
    side: typeof sides.top,
    x0: number,
    y0: number,
    x1: number,
    y1: number,
    doubleInside: 1 | -1,
  ): void => {
    if (!side) return;
    const cfg = typeof side === 'object' ? side : { style: 'thin' as const };
    const color =
      (typeof side === 'object' && side.color) || theme.fgStrong || theme.fg || '#000000';
    const widthMap: Record<string, number> = {
      thin: 1,
      medium: 1.6,
      thick: 2.5,
      dashed: 1,
      dotted: 1,
      double: 1,
      hair: 0.5,
      mediumDashed: 1.6,
      dashDot: 1,
      mediumDashDot: 1.6,
      dashDotDot: 1,
      mediumDashDotDot: 1.6,
      slantDashDot: 1,
    };
    const dashMap: Record<string, number[]> = {
      thin: [],
      medium: [],
      thick: [],
      dashed: [4, 3],
      dotted: [1, 2],
      double: [],
      hair: [1, 1],
      mediumDashed: [6, 3],
      dashDot: [4, 2, 1, 2],
      mediumDashDot: [6, 2, 2, 2],
      dashDotDot: [4, 2, 1, 2, 1, 2],
      mediumDashDotDot: [6, 2, 2, 2, 2, 2],
      slantDashDot: [5, 2, 1, 2],
    };
    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = widthMap[cfg.style] ?? 1;
    ctx.setLineDash(dashMap[cfg.style] ?? []);
    if (cfg.style === 'double') {
      // Spreadsheets keep double borders inside the owning cell. Drawing both
      // strokes inward avoids clipping on viewport and sheet edges.
      const gap = 3;
      const horizontal = y0 === y1;
      const dx = horizontal ? 0 : doubleInside * gap;
      const dy = horizontal ? doubleInside * gap : 0;
      ctx.beginPath();
      ctx.moveTo(x0, y0);
      ctx.lineTo(x1, y1);
      ctx.moveTo(x0 + dx, y0 + dy);
      ctx.lineTo(x1 + dx, y1 + dy);
      ctx.stroke();
    } else {
      ctx.beginPath();
      ctx.moveTo(x0, y0);
      ctx.lineTo(x1, y1);
      ctx.stroke();
    }
    ctx.restore();
  };
  const yt = Math.round(bounds.y) + 0.5;
  const yb = Math.round(bounds.y + bounds.h) - 0.5;
  const xl = Math.round(bounds.x) + 0.5;
  const xr = Math.round(bounds.x + bounds.w) - 0.5;
  // Keep user borders inside the cell rect. This is especially important for
  // row/column 1 and viewport edges, where centered strokes can be clipped.
  drawSide(sides.top, xl, yt, xr, yt, 1);
  drawSide(sides.bottom, xl, yb, xr, yb, -1);
  drawSide(sides.left, xl, yt, xl, yb, 1);
  drawSide(sides.right, xr, yt, xr, yb, -1);
  drawSide(sides.diagonalDown, xl, yt, xr, yb, 1);
  drawSide(sides.diagonalUp, xl, yb, xr, yt, 1);
}
