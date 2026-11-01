import type {
  CellBorderSide,
  CellBorderStyle,
  CellGradientFill,
  CellVAlign,
  ConditionalIconSet,
} from '../../store/store.js';
import type { Rect } from '../geometry.js';
import { fillRect, surfaceRect } from '../grid/chrome-context.js';
import { drawConditionalIconArt } from './cf-icons.js';
import { visualZoom } from './text.js';
import type { CellPaintCtx } from './types.js';

/* ---- Base painters wired in MS-A. Future slots (formatting, CF, DV,
 * comments, hyperlinks) live next to these, capability-gated. ---- */

export function paintCellBackground({
  ctx,
  bounds,
  theme,
  isInRange,
  isActive,
  dpr,
  rtl,
}: CellPaintCtx): void {
  if (isActive) {
    // Only the surface clears: the rules the cell owns keep the range shading.
    const surface = surfaceRect(bounds, dpr ?? 1, rtl === true);
    ctx.fillStyle = theme.bgElev;
    ctx.fillRect(surface.x, surface.y, surface.w, surface.h);
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

/** Icon box edge, its inset from the leading edge and its clearance above
 *  the cell's own bottom gridline, in CSS px at 100% zoom. */
const ICON_BOX = 11;
const ICON_INSET = 1;
const ICON_BOTTOM_GAP = 2;

/** Paint a small icon-set glyph at the leading edge of a cell rect. The
 *  artwork is drawn with primitive shapes — no SVG / image asset — so the
 *  renderer stays self-contained. `slot` is 0-based; the family decides the
 *  count and how the slot maps to colour/direction. The glyph follows the
 *  cell's vertical alignment the way its text does (bottom by default) and
 *  scales with `zoom`. */
export function paintConditionalIcon(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  kind: ConditionalIconSet,
  slot: number,
  rtl = false,
  vAlign?: CellVAlign,
  zoom?: number,
): void {
  const z = visualZoom(zoom);
  const size = ICON_BOX * z;
  // The last CSS px row of `rect` is the cell's own gridline.
  const bottom = rect.y + rect.h - 1 - ICON_BOTTOM_GAP * z;
  const y =
    vAlign === 'top'
      ? rect.y
      : vAlign === undefined || vAlign === 'bottom'
        ? bottom - size
        : (rect.y + bottom - size) / 2;
  const x = rtl ? rect.x + rect.w - (ICON_INSET + ICON_BOX) * z : rect.x + ICON_INSET * z;
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.w, rect.h);
  ctx.clip();
  ctx.translate(x, y);
  ctx.scale(z, z);
  drawConditionalIconArt(ctx, kind, slot);
  ctx.restore();
}

/** OOXML gradient fill. A linear gradient's `degree` turns clockwise from
 *  left-to-right; a path gradient is approximated by a radial one spreading
 *  from the centre of its inner rectangle to the farthest cell corner. */
function paintGradientFill(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  gradient: CellGradientFill,
): void {
  let paint: CanvasGradient;
  if (gradient.type === 'path') {
    const cx = bounds.x + ((gradient.left + gradient.right) / 2) * bounds.w;
    const cy = bounds.y + ((gradient.top + gradient.bottom) / 2) * bounds.h;
    const radius = Math.max(
      Math.hypot(cx - bounds.x, cy - bounds.y),
      Math.hypot(bounds.x + bounds.w - cx, cy - bounds.y),
      Math.hypot(cx - bounds.x, bounds.y + bounds.h - cy),
      Math.hypot(bounds.x + bounds.w - cx, bounds.y + bounds.h - cy),
    );
    paint = ctx.createRadialGradient(cx, cy, 0, cx, cy, radius);
  } else {
    const angle = (gradient.degree * Math.PI) / 180;
    const dx = Math.cos(angle);
    const dy = Math.sin(angle);
    const half = (Math.abs(bounds.w * dx) + Math.abs(bounds.h * dy)) / 2;
    const cx = bounds.x + bounds.w / 2;
    const cy = bounds.y + bounds.h / 2;
    paint = ctx.createLinearGradient(
      cx - dx * half,
      cy - dy * half,
      cx + dx * half,
      cy + dy * half,
    );
  }
  for (const stop of gradient.stops) {
    paint.addColorStop(Math.max(0, Math.min(1, stop.position)), stop.color);
  }
  ctx.save();
  ctx.fillStyle = paint;
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  ctx.restore();
}

/** User-set background fill — drawn above range tint and active highlight so
 *  cells with explicit fill stay visually filled even when selected. */
export function paintCellFill({ ctx, bounds: cell, format, dpr, rtl }: CellPaintCtx): void {
  // A fill covers the gridline before the cell as well as its own.
  const bounds = fillRect(cell, dpr ?? 1, rtl === true);
  if (format?.fillGradient) {
    paintGradientFill(ctx, bounds, format.fillGradient);
    return;
  }
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
    case 'gray0625':
    case 'gray125':
    case 'gray25':
    case 'gray50':
    case 'gray75': {
      const step =
        pattern === 'gray0625'
          ? 8
          : pattern === 'gray125'
            ? 6
            : pattern === 'gray25'
              ? 4
              : pattern === 'gray50'
                ? 3
                : 2;
      const size = pattern === 'gray50' ? 1.5 : pattern === 'gray75' ? 2 : 1;
      for (let y = bounds.y + 2; y < bounds.y + bounds.h; y += step) {
        for (let x = bounds.x + 2; x < bounds.x + bounds.w; x += step) {
          ctx.fillRect(x, y, size, size);
        }
      }
      break;
    }
    case 'horizontal':
    case 'darkHorizontal':
    case 'lightHorizontal': {
      const step = pattern === 'lightHorizontal' ? 7 : 4;
      for (let y = bounds.y + step - 1; y < bounds.y + bounds.h; y += step) {
        ctx.beginPath();
        ctx.moveTo(bounds.x, y);
        ctx.lineTo(bounds.x + bounds.w, y);
        ctx.stroke();
      }
      break;
    }
    case 'vertical':
    case 'darkVertical':
    case 'lightVertical': {
      const step = pattern === 'lightVertical' ? 7 : 4;
      for (let x = bounds.x + step - 1; x < bounds.x + bounds.w; x += step) {
        ctx.beginPath();
        ctx.moveTo(x, bounds.y);
        ctx.lineTo(x, bounds.y + bounds.h);
        ctx.stroke();
      }
      break;
    }
    case 'diagonalDown':
    case 'diagonalUp':
    case 'darkDown':
    case 'darkUp':
    case 'lightDown':
    case 'lightUp': {
      const direction =
        pattern === 'diagonalDown' || pattern === 'darkDown' || pattern === 'lightDown' ? 1 : -1;
      const step = pattern === 'lightDown' || pattern === 'lightUp' ? 9 : 6;
      for (let x = bounds.x - bounds.h; x < bounds.x + bounds.w; x += step) {
        ctx.beginPath();
        ctx.moveTo(x, direction > 0 ? bounds.y : bounds.y + bounds.h);
        ctx.lineTo(x + bounds.h, direction > 0 ? bounds.y + bounds.h : bounds.y);
        ctx.stroke();
      }
      break;
    }
    case 'darkGrid':
    case 'lightGrid': {
      const step = pattern === 'lightGrid' ? 7 : 4;
      for (let y = bounds.y + step - 1; y < bounds.y + bounds.h; y += step) {
        ctx.beginPath();
        ctx.moveTo(bounds.x, y);
        ctx.lineTo(bounds.x + bounds.w, y);
        ctx.stroke();
      }
      for (let x = bounds.x + step - 1; x < bounds.x + bounds.w; x += step) {
        ctx.beginPath();
        ctx.moveTo(x, bounds.y);
        ctx.lineTo(x, bounds.y + bounds.h);
        ctx.stroke();
      }
      break;
    }
    case 'darkTrellis':
    case 'lightTrellis': {
      const step = pattern === 'lightTrellis' ? 9 : 6;
      for (const direction of [1, -1]) {
        for (let x = bounds.x - bounds.h; x < bounds.x + bounds.w; x += step) {
          ctx.beginPath();
          ctx.moveTo(x, direction > 0 ? bounds.y : bounds.y + bounds.h);
          ctx.lineTo(x + bounds.h, direction > 0 ? bounds.y + bounds.h : bounds.y);
          ctx.stroke();
        }
      }
      break;
    }
  }
  ctx.restore();
}

/** Perpendicular width of a diagonal border, in CSS px at one rule per px. */
const DIAGONAL_WIDTH: Partial<Record<CellBorderStyle, number>> = {
  medium: 2.67,
  mediumDashed: 2.67,
  mediumDashDot: 2.67,
  mediumDashDotDot: 2.67,
  slantDashDot: 2.67,
  thick: 3.46,
};

/** Diagonal dash runs in CSS px, measured along x as the spreadsheet does. */
const DIAGONAL_DASH: Partial<Record<CellBorderStyle, readonly number[]>> = {
  dashed: [2, 2],
  dotted: [1, 1, 1, 1, 1, 5],
  mediumDashed: [9, 3],
  dashDot: [9, 3, 3, 3],
  mediumDashDot: [9, 3, 3, 3],
  dashDotDot: [9, 3, 3, 3, 3, 3],
  mediumDashDotDot: [9, 3, 3, 3, 3, 3],
};

/** Diagonal border lines across the cell surface. They run between the
 *  centres of the corner gridlines and are clipped to the fill rectangle, so
 *  wide strokes end square. Edge borders are painted per edge by
 *  `render/grid/borders.ts`. */
export function paintCellDiagonals({ ctx, bounds, theme, format, dpr = 1 }: CellPaintCtx): void {
  const sides = format?.borders;
  if (!sides?.diagonalDown && !sides?.diagonalUp) return;
  const g = Math.max(1, Math.round(dpr));
  const rule = g / dpr;
  const l = Math.round(bounds.x * dpr) / dpr;
  const r = Math.round((bounds.x + bounds.w) * dpr) / dpr;
  const t = Math.round(bounds.y * dpr) / dpr;
  const b = Math.round((bounds.y + bounds.h) * dpr) / dpr;
  const x0 = l - rule / 2;
  const x1 = r - rule / 2;
  ctx.save();
  ctx.beginPath();
  ctx.rect(l - rule, t - rule, r - l + rule, b - t + rule);
  ctx.clip();
  const draw = (side: CellBorderSide | undefined, ya: number, yb: number): void => {
    if (!side) return;
    const style = typeof side === 'object' ? side.style : 'thin';
    ctx.strokeStyle = (typeof side === 'object' && side.color) || theme.cellAutoFg;
    ctx.lineWidth = (DIAGONAL_WIDTH[style] ?? 1) * rule;
    const stretch = Math.hypot(x1 - x0, yb - ya) / Math.max(x1 - x0, 1);
    ctx.setLineDash((DIAGONAL_DASH[style] ?? []).map((d) => d * rule * stretch));
    ctx.beginPath();
    // A double diagonal is two single lines one rule above and below.
    for (const dy of style === 'double' ? [-rule, rule] : [0]) {
      ctx.moveTo(x0, ya + dy);
      ctx.lineTo(x1, yb + dy);
    }
    ctx.stroke();
  };
  draw(sides.diagonalDown, t - rule / 2, b - rule / 2);
  draw(sides.diagonalUp, b - rule / 2, t - rule / 2);
  ctx.restore();
}
