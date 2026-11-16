import type {
  CellBorderSide,
  CellBorderStyle,
  CellFormat,
  CellGradientFill,
  CellVAlign,
  ConditionalIconSet,
} from '../../store/store.js';
import type { Rect } from '../geometry.js';
import { fillRect, rulePx, surfaceRect } from '../grid/chrome-context.js';
import { drawConditionalIconArt } from './cf-icons.js';
import { visualZoom } from './text.js';
import type { CellPaintCtx } from './types.js';

/* ---- Base cell painters: background, fill, borders, diagonals, text. ---- */

/** Whether the format carries its own background: solid, pattern or gradient. */
export const hasCellFill = (fmt: Partial<CellFormat> | undefined): boolean =>
  !!(fmt?.fill || fmt?.fillPattern || fmt?.fillGradient);

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
 *  left-to-right. A path gradient is a box: each of the four triangles between
 *  the centre of its inner rectangle and the cell corners ramps linearly from
 *  the centre to the edge it faces. */
function paintGradientFill(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  gradient: CellGradientFill,
): void {
  const addStops = (paint: CanvasGradient): CanvasGradient => {
    for (const stop of gradient.stops) {
      paint.addColorStop(Math.max(0, Math.min(1, stop.position)), stop.color);
    }
    return paint;
  };
  ctx.save();
  if (gradient.type === 'path') {
    const l = bounds.x;
    const t = bounds.y;
    const r = bounds.x + bounds.w;
    const b = bounds.y + bounds.h;
    const cx = l + ((gradient.left + gradient.right) / 2) * bounds.w;
    const cy = t + ((gradient.top + gradient.bottom) / 2) * bounds.h;
    const sides: [number, number, number, number, number, number][] = [
      [l, t, r, t, cx, t],
      [r, t, r, b, r, cy],
      [r, b, l, b, cx, b],
      [l, b, l, t, l, cy],
    ];
    // Each wedge reaches a pixel past its seams, so antialiased clip edges
    // never leave the background showing between two wedges.
    const bleed = 1;
    for (const [ax, ay, bx, by, ex, ey] of sides) {
      const ux = Math.sign(bx - ax) * bleed;
      const uy = Math.sign(by - ay) * bleed;
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.lineTo(ax - ux, ay - uy);
      ctx.lineTo(bx + ux, by + uy);
      ctx.closePath();
      ctx.clip();
      ctx.fillStyle = addStops(ctx.createLinearGradient(cx, cy, ex, ey));
      ctx.fillRect(l, t, bounds.w, bounds.h);
      ctx.restore();
    }
    ctx.restore();
    return;
  }
  const angle = (gradient.degree * Math.PI) / 180;
  const dx = Math.cos(angle);
  const dy = Math.sin(angle);
  const half = (Math.abs(bounds.w * dx) + Math.abs(bounds.h * dy)) / 2;
  const cx = bounds.x + bounds.w / 2;
  const cy = bounds.y + bounds.h / 2;
  ctx.fillStyle = addStops(
    ctx.createLinearGradient(cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half),
  );
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  ctx.restore();
}

/** Excel's 8x8 pattern bitmaps (X = foreground), one device pixel per bitmap
 *  pixel, anchored at sheet pixel (0,0). Rows repeat to fill all eight. */
const PATTERN_TILES: Record<string, readonly string[]> = {
  mediumGray: ['X.X.X.X.', '.X.X.X.X'],
  darkGray: ['X.XXX.XX', 'XXX.XXX.'],
  lightGray: ['..X...X.', 'X...X...'],
  darkHorizontal: ['........', '........', 'XXXXXXXX', 'XXXXXXXX'],
  darkVertical: ['..XX..XX'],
  darkDown: ['XX..XX..', '.XX..XX.', '..XX..XX', 'X..XX..X'],
  darkUp: ['..XX..XX', '.XX..XX.', 'XX..XX..', 'X..XX..X'],
  darkGrid: ['X..XX..X', '.XX..XX.', '.XX..XX.', 'X..XX..X'],
  darkTrellis: ['XXXXXXXX', '.XX..XX.', 'XXXXXXXX', 'X..XX..X'],
  lightHorizontal: ['........', '........', 'XXXXXXXX', '........'],
  lightVertical: ['..X...X.'],
  lightDown: ['X...X...', '.X...X..', '..X...X.', '...X...X'],
  lightUp: ['...X...X', '..X...X.', '.X...X..', 'X...X...'],
  lightGrid: ['..X...X.', '..X...X.', 'XXXXXXXX', '..X...X.'],
  lightTrellis: ['X...X...', '.X.X.X.X', '..X...X.', '.X.X.X.X'],
  gray125: ['X...X...', '........', '..X...X.', '........'],
  gray0625: ['......X.', '........', '..X.....', '........'],
};

/** Spreadsheet-format pattern names that share another pattern's bitmap. */
const PATTERN_ALIAS: Record<string, string> = {
  gray50: 'mediumGray',
  gray75: 'darkGray',
  gray25: 'lightGray',
  horizontal: 'darkHorizontal',
  vertical: 'darkVertical',
  diagonalDown: 'darkDown',
  diagonalUp: 'darkUp',
};

/** Foreground runs `[start, length]` along one 8-pixel tile row. */
const tileRuns = (row: string): [number, number][] => {
  const runs: [number, number][] = [];
  for (let x = 0; x < 8; x++) {
    if (row[x] !== 'X') continue;
    const last = runs[runs.length - 1];
    if (last && last[0] + last[1] === x) last[1]++;
    else runs.push([x, 1]);
  }
  return runs;
};

const tileRunCache = new WeakMap<readonly string[], [number, number][][]>();

const tileRunsOf = (tile: readonly string[]): [number, number][][] => {
  let runs = tileRunCache.get(tile);
  if (!runs) {
    runs = tile.map(tileRuns);
    tileRunCache.set(tile, runs);
  }
  return runs;
};

/** Pattern ink over `bounds`, drawn per device pixel from the sheet-anchored
 *  tile so the phase does not depend on the cell the pattern sits in. */
function paintPatternTile(
  ctx: CanvasRenderingContext2D,
  bounds: Rect,
  tile: readonly string[],
  color: string,
  dpr: number,
): void {
  const x0 = Math.round(bounds.x * dpr);
  const y0 = Math.round(bounds.y * dpr);
  const x1 = Math.round((bounds.x + bounds.w) * dpr);
  const y1 = Math.round((bounds.y + bounds.h) * dpr);
  const runs = tileRunsOf(tile);
  ctx.fillStyle = color;
  ctx.beginPath();
  for (let y = y0; y < y1; y++) {
    const rowRuns = runs[y % tile.length] ?? [];
    if (rowRuns.length === 0) continue;
    for (let base = x0 - (((x0 % 8) + 8) % 8); base < x1; base += 8) {
      for (const [start, len] of rowRuns) {
        const a = Math.max(base + start, x0);
        const b = Math.min(base + start + len, x1);
        if (b > a) ctx.rect(a / dpr, y / dpr, (b - a) / dpr, 1 / dpr);
      }
    }
  }
  ctx.fill();
}

/** User-set background fill — drawn above range tint and active highlight so
 *  cells with explicit fill stay visually filled even when selected. */
export function paintCellFill({ ctx, bounds: cell, theme, format, dpr, rtl }: CellPaintCtx): void {
  // A fill covers the gridline before the cell as well as its own.
  const bounds = fillRect(cell, dpr ?? 1, rtl === true);
  if (format?.fillGradient) {
    paintGradientFill(ctx, bounds, format.fillGradient);
    return;
  }
  const fill = format?.fill;
  const pattern = format?.fillPattern;
  if (!fill && !pattern) return;
  // A pattern without a background still replaces what lay under the cell,
  // neighbours' fills on the shared gridline included.
  ctx.fillStyle = fill ?? theme.bg;
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  if (!pattern) return;
  const tile = PATTERN_TILES[PATTERN_ALIAS[pattern] ?? pattern];
  if (!tile) return;
  ctx.save();
  paintPatternTile(ctx, bounds, tile, format?.fillPatternColor ?? '#000000', dpr ?? 1);
  ctx.restore();
}

/** Selected cells other than the active one darken their fill by 0.776. */
export function paintSelectedFillShade({ ctx, bounds: cell, dpr, rtl }: CellPaintCtx): void {
  const bounds = fillRect(cell, dpr ?? 1, rtl === true);
  ctx.fillStyle = 'rgba(0, 0, 0, 0.224)';
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
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

/** A rule-wide diagonal sits this many device px right (down-diagonal) or left
 *  (up-diagonal) of the corner-gridline centres and is this fraction of a rule wide (Skia coverage matches Excel ink to 0.5%);
 *  wider ones do not. Excel antialiases analytically, so edge pixels still
 *  differ by a few levels. */
const RULE_DIAGONAL_SHIFT = 1.41;
const RULE_DIAGONAL_INK = 0.95;

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
  const g = rulePx(dpr);
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
  const draw = (side: CellBorderSide | undefined, ya: number, yb: number, dir = 1): void => {
    if (!side) return;
    const style = typeof side === 'object' ? side.style : 'thin';
    ctx.strokeStyle = (typeof side === 'object' && side.color) || theme.cellAutoFg;
    const wide = DIAGONAL_WIDTH[style] !== undefined;
    ctx.lineWidth = (DIAGONAL_WIDTH[style] ?? RULE_DIAGONAL_INK) * rule;
    const shift = wide ? 0 : (dir * RULE_DIAGONAL_SHIFT) / dpr;
    const stretch = Math.hypot(x1 - x0, yb - ya) / Math.max(x1 - x0, 1);
    ctx.setLineDash((DIAGONAL_DASH[style] ?? []).map((d) => d * rule * stretch));
    ctx.beginPath();
    // A double diagonal is two single lines one rule above and below.
    for (const dy of style === 'double' ? [-rule, rule] : [0]) {
      ctx.moveTo(x0 + shift, ya + dy);
      ctx.lineTo(x1 + shift, yb + dy);
    }
    ctx.stroke();
  };
  draw(sides.diagonalDown, t - rule / 2, b - rule / 2);
  draw(sides.diagonalUp, b - rule / 2, t - rule / 2, -1);
  ctx.restore();
}
