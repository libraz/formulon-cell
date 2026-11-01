import type {
  CellBorderSide,
  CellBorderStyle,
  CellGradientFill,
  ConditionalIconSet,
} from '../../store/store.js';
import type { Rect } from '../geometry.js';
import { fillRect } from '../grid/chrome-context.js';
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

const ARROWS_3 = [Math.PI / 2, 0, -Math.PI / 2];
const ARROWS_4 = [Math.PI / 2, Math.PI / 4, -Math.PI / 4, -Math.PI / 2];
const ARROWS_5 = [Math.PI / 2, Math.PI / 4, 0, -Math.PI / 4, -Math.PI / 2];

/** Arrow directions per slot, low → high (0 = right, π/2 = down). */
const ARROW_ANGLES: Partial<Record<ConditionalIconSet, readonly number[]>> = {
  arrows3: ARROWS_3,
  arrowsGray3: ARROWS_3,
  arrows4: ARROWS_4,
  arrowsGray4: ARROWS_4,
  arrows5: ARROWS_5,
  arrowsGray5: ARROWS_5,
};

const RAMP: Readonly<Record<number, readonly string[]>> = {
  3: ['#d24545', '#cfa64a', '#3aa055'],
  4: ['#d24545', '#cfa64a', '#cfa64a', '#3aa055'],
  5: ['#d24545', '#e07a4d', '#cfa64a', '#7fb352', '#3aa055'],
};

const rampColor = (slots: number, idx: number): string | undefined => RAMP[slots]?.[idx];

/** Paint a small icon-set glyph at the left of a cell rect. The artwork is
 *  drawn with primitive shapes — no SVG / image asset — so the renderer
 *  stays self-contained. `slot` is 0-based; the family decides the count
 *  (3 or 5) and how the slot maps to color/direction. */
export function paintConditionalIcon(
  ctx: CanvasRenderingContext2D,
  rect: Rect,
  kind: ConditionalIconSet,
  slot: number,
  rtl = false,
): void {
  const cx = rtl
    ? rect.x + rect.w - CONDITIONAL_ICON_GUTTER / 2
    : rect.x + CONDITIONAL_ICON_GUTTER / 2;
  const cy = rect.y + rect.h / 2;
  const r = Math.min(5, Math.floor(rect.h * 0.32));
  ctx.save();
  if (ARROW_ANGLES[kind]) {
    // Direction tracks the slot (low → down, high → up); colour runs
    // red → amber → green, or stays gray for the gray families.
    const angles = ARROW_ANGLES[kind] ?? [];
    const idx = Math.max(0, Math.min(angles.length - 1, slot));
    const color = kind.startsWith('arrowsGray')
      ? '#7f7f7f'
      : (rampColor(angles.length, idx) ?? '#777');
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
  } else if (
    kind === 'traffic3' ||
    kind === 'trafficRim3' ||
    kind === 'traffic4' ||
    kind === 'redToBlack4'
  ) {
    // One solid circle in the slot's colour; the rimmed set adds a housing.
    const colors =
      kind === 'traffic4'
        ? ['#404040', '#d24545', '#cfa64a', '#3aa055']
        : kind === 'redToBlack4'
          ? ['#404040', '#8c8c8c', '#e8a0a0', '#d24545']
          : ['#d24545', '#cfa64a', '#3aa055'];
    const idx = Math.max(0, Math.min(colors.length - 1, slot));
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
  } else if (kind === 'signs3') {
    // Red diamond, amber triangle, green circle.
    const idx = Math.max(0, Math.min(2, slot));
    ctx.fillStyle = ['#d24545', '#cfa64a', '#3aa055'][idx] ?? '#777';
    ctx.beginPath();
    if (idx === 0) {
      ctx.moveTo(cx, cy - r);
      ctx.lineTo(cx + r, cy);
      ctx.lineTo(cx, cy + r);
      ctx.lineTo(cx - r, cy);
      ctx.closePath();
    } else if (idx === 1) {
      ctx.moveTo(cx, cy - r);
      ctx.lineTo(cx + r, cy + r);
      ctx.lineTo(cx - r, cy + r);
      ctx.closePath();
    } else {
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
    }
    ctx.fill();
  } else if (kind === 'symbols3' || kind === 'symbolsCircled3') {
    const idx = Math.max(0, Math.min(2, slot));
    ctx.lineWidth = 1.6;
    if (kind === 'symbolsCircled3') {
      // Coloured disc with the symbol knocked out in white.
      ctx.fillStyle = ['#d24545', '#cfa64a', '#3aa055'][idx] ?? '#777';
      ctx.beginPath();
      ctx.arc(cx, cy, r + 1, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#ffffff';
      ctx.fillStyle = '#ffffff';
      ctx.beginPath();
      if (idx === 2) {
        ctx.moveTo(cx - r + 2, cy);
        ctx.lineTo(cx - 1, cy + r - 2);
        ctx.lineTo(cx + r - 2, cy - r + 2);
      } else if (idx === 1) {
        ctx.moveTo(cx, cy - r + 2);
        ctx.lineTo(cx, cy + 1);
      } else {
        ctx.moveTo(cx - r + 2, cy - r + 2);
        ctx.lineTo(cx + r - 2, cy + r - 2);
        ctx.moveTo(cx + r - 2, cy - r + 2);
        ctx.lineTo(cx - r + 2, cy + r - 2);
      }
      ctx.stroke();
      ctx.restore();
      return;
    }
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
  } else if (kind === 'quarters5') {
    const idx = Math.max(0, Math.min(4, slot));
    ctx.strokeStyle = '#5c5c5c';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.stroke();
    ctx.fillStyle = '#3f7ad8';
    if (idx > 0) {
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + (idx / 4) * Math.PI * 2);
      ctx.closePath();
      ctx.fill();
    }
  } else if (kind === 'ratings4' || kind === 'ratings5') {
    // Four signal bars; the slot sets how many are lit.
    const lit =
      kind === 'ratings4' ? Math.max(0, Math.min(3, slot)) + 1 : Math.max(0, Math.min(4, slot));
    ctx.fillStyle = '#3f7ad8';
    const baseX = cx - 6;
    for (let i = 0; i < 4; i += 1) {
      const h = 3 + i * 2.5;
      ctx.globalAlpha = i < lit ? 1 : 0.25;
      ctx.fillRect(baseX + i * 3.2, cy + r - h, 2.2, h);
    }
    ctx.globalAlpha = 1;
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
