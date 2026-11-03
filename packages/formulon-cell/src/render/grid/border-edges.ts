// Cell borders as the spreadsheet paints them, resolved per edge and drawn in
// whole device pixels. All offsets are relative to a boundary E, where the
// gridline occupies [E - g, E) and g is one CSS pixel snapped to device pixels.
//
// - A side shared by two cells shows one winner: the heavier style, and the
//   later cell (below / to the right) when the styles are identical.
// - Thin lines sit on the gridline, medium lines grow one gridline width into
//   the earlier cell, thick and double straddle the gridline.
// - A lone segment spans [L - g, R); at a corner it reaches the far edge of a
//   perpendicular border, and double lines join outer-to-outer, inner-to-inner.
// - Dash phases are anchored to the canvas origin so runs stay continuous
//   across cells. A bordered edge replaces the gridline beneath its gaps.

import type { CellBorderSide, CellBorderStyle } from '../../store/store.js';

export interface BorderEdge {
  readonly style: CellBorderStyle;
  readonly color: string | undefined;
}

/** Conflict precedence, lowest first. Weight classes rank first; inside a
 *  class the solid line beats the dashed variants in this order. */
const PRECEDENCE: readonly CellBorderStyle[] = [
  'hair',
  'dashDotDot',
  'dashDot',
  'dotted',
  'dashed',
  'thin',
  'mediumDashDotDot',
  'slantDashDot',
  'mediumDashDot',
  'mediumDashed',
  'medium',
  'thick',
  'double',
];
const RANK = new Map(PRECEDENCE.map((style, i) => [style, i]));

export const borderEdge = (side: CellBorderSide | undefined): BorderEdge | null => {
  if (!side) return null;
  return typeof side === 'object'
    ? { style: side.style, color: side.color }
    : { style: 'thin', color: undefined };
};

/** The edge shown where `earlier` (above / left) meets `later`. */
export const resolveBorderEdge = (
  earlier: BorderEdge | null,
  later: BorderEdge | null,
): BorderEdge | null => {
  if (!earlier) return later;
  if (!later) return earlier;
  return (RANK.get(earlier.style) ?? 0) > (RANK.get(later.style) ?? 0) ? earlier : later;
};

type Stroke = readonly [start: number, end: number];

/** Across-axis strokes of `style` relative to the boundary, in device px. */
export const borderStrokes = (style: CellBorderStyle, g: number): readonly Stroke[] => {
  switch (style) {
    case 'thick':
      return [[-2 * g, g]];
    case 'double':
      return [
        [-2 * g, -g],
        [0, g],
      ];
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return [[-2 * g, 0]];
    default:
      return [[-g, 0]];
  }
};

const extentOf = (strokes: readonly Stroke[]): Stroke => [
  Math.min(...strokes.map((s) => s[0])),
  Math.max(...strokes.map((s) => s[1])),
];

/** On/off dash lengths in CSS px, scaled by g and repeated from the origin. */
const DASHES: Partial<Record<CellBorderStyle, readonly number[]>> = {
  mediumDashed: [9, 3],
  dashDot: [9, 3, 3, 3],
  mediumDashDot: [9, 3, 3, 3],
  dashDotDot: [9, 3, 3, 3, 3, 3],
  mediumDashDotDot: [9, 3, 3, 3, 3, 3],
};

const mod = (n: number, m: number): number => ((n % m) + m) % m;

/** Ink coverage across a slant cut, which the spreadsheet antialiases. */
const SLANT_CUT = [0.855, 0.376, 0.376, 0.855];

/** Ink coverage (0..1) of the device pixel at `along` / `across`; null for a
 *  solid style. `along` is relative to the grid origin. `across` is relative
 *  to the grid origin for the hairline family, which staggers between pixel
 *  rows, and to the stroke's first pixel row for slant dash-dot, whose
 *  antialiased diagonal cuts restart on every line. */
export const borderPixelInk = (
  style: CellBorderStyle,
  along: number,
  across: number,
  g: number,
): number | null => {
  if (style !== 'slantDashDot') {
    const on = borderPixelOn(style, along, across, g);
    return on === null ? null : on ? 1 : 0;
  }
  const at = mod(along + 2 * across, 18 * g);
  for (const cut of [12.5, 0.5]) {
    const offset = at - Math.round(cut * g);
    if (offset >= 0 && offset < SLANT_CUT.length) return SLANT_CUT[offset] ?? 1;
  }
  return 1;
};

/** Whether the device pixel at `along` / `across` (relative to the grid
 *  origin) is inked, or null for a solid or slant style. */
export const borderPixelOn = (
  style: CellBorderStyle,
  along: number,
  across: number,
  g: number,
): boolean | null => {
  switch (style) {
    case 'hair':
      return mod(along + across, 2) === 0;
    case 'dashed':
      return mod(along + across, 4) !== 3;
    case 'dotted':
      return mod(Math.floor((along - 1) / 2) + across, 2) === 0;
    default: {
      const dashes = DASHES[style];
      if (!dashes) return null;
      const period = dashes.reduce((sum, d) => sum + d, 0) * g;
      let at = mod(along, period);
      for (let i = 0; i < dashes.length; i += 1) {
        const len = (dashes[i] ?? 0) * g;
        if (at < len) return i % 2 === 0;
        at -= len;
      }
      return false;
    }
  }
};

/** Along-axis span of each stroke of an edge between boundaries `a0` and
 *  `a1` (device px). `before` / `after` are the perpendicular edges meeting
 *  each end on the earlier and later side of this edge's axis; `continued`
 *  says the same line already runs up to `a0` from the previous cell. */
export const strokeSpans = (
  edge: BorderEdge,
  g: number,
  a0: number,
  a1: number,
  start: { before: BorderEdge | null; after: BorderEdge | null },
  end: { before: BorderEdge | null; after: BorderEdge | null },
  continued = false,
): readonly Stroke[] => {
  const strokes = borderStrokes(edge.style, g);
  const reach = (
    perps: { before: BorderEdge | null; after: BorderEdge | null },
    atStart: boolean,
    strokeIndex: number,
  ): number => {
    const base = atStart ? (continued ? a0 : a0 - g) : a1;
    const at = atStart ? a0 : a1;
    const doubleBefore = perps.before?.style === 'double';
    const doubleAfter = perps.after?.style === 'double';
    if (edge.style === 'double' && doubleBefore !== doubleAfter) {
      // The stroke on the side the perpendicular double leaves from is the
      // outer one of the corner; the other stroke stops at the inner line.
      const outer = doubleAfter ? strokeIndex === 0 : strokeIndex === 1;
      if (atStart) return outer ? at - 2 * g : at;
      return outer ? at + g : at - g;
    }
    // Lines run through a corner, except that a dashed line stops short of
    // the perpendicular band; `dashCorner` inks the corner square itself.
    const stopsShort = DASHES[edge.style] !== undefined;
    const ends = [perps.before, perps.after]
      .filter((perp): perp is BorderEdge => perp !== null)
      .map((perp) => extentOf(borderStrokes(perp.style, g)));
    if (ends.length === 0) return base;
    if (stopsShort) {
      return atStart
        ? at + Math.max(...ends.map((x) => x[1]))
        : at + Math.min(...ends.map((x) => x[0]));
    }
    return atStart
      ? Math.min(base, at + Math.min(...ends.map((x) => x[0])))
      : Math.max(base, at + Math.max(...ends.map((x) => x[1])));
  };
  return strokes.map((_, i) => [reach(start, true, i), reach(end, false, i)] as const);
};

export const isDashed = (style: CellBorderStyle): boolean => DASHES[style] !== undefined;

/** Pixels of the corner square where a dashed horizontal meets a dashed
 *  vertical: inked only where both dash patterns are on. Coordinates are
 *  device px relative to the grid origin; `x0..x1` is the vertical band and
 *  `y0..y1` the horizontal one. */
export const dashCornerPixels = (
  horizontal: CellBorderStyle,
  vertical: CellBorderStyle,
  g: number,
  x0: number,
  x1: number,
  y0: number,
  y1: number,
): { x: number; y: number }[] => {
  const out: { x: number; y: number }[] = [];
  for (let y = y0; y < y1; y += 1) {
    if (!borderPixelOn(vertical, y, 0, g)) continue;
    for (let x = x0; x < x1; x += 1) {
      if (borderPixelOn(horizontal, x, 0, g)) out.push({ x, y });
    }
  }
  return out;
};

export interface EdgeRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** Device-pixel rectangles of every stroke of an edge on boundary `e`. */
export const edgeRects = (
  edge: BorderEdge,
  horizontal: boolean,
  e: number,
  spans: readonly Stroke[],
  g: number,
): EdgeRect[] =>
  borderStrokes(edge.style, g).map(([s0, s1], i) => {
    const [a0, a1] = spans[i] ?? spans[0] ?? [0, 0];
    return horizontal
      ? { x: a0, y: e + s0, w: Math.max(0, a1 - a0), h: s1 - s0 }
      : { x: e + s0, y: a0, w: s1 - s0, h: Math.max(0, a1 - a0) };
  });

const tiles = new Map<string, CanvasPattern | null>();

/** A repeating tile of the dash pattern, anchored at the grid origin. */
const patternTile = (
  ctx: CanvasRenderingContext2D,
  style: CellBorderStyle,
  horizontal: boolean,
  color: string,
  g: number,
): CanvasPattern | null => {
  const key = `${style}|${horizontal ? 'h' : 'v'}|${color}|${g}`;
  if (tiles.has(key)) return tiles.get(key) ?? null;
  const doc = ctx.canvas?.ownerDocument;
  const dashes = DASHES[style];
  const along =
    style === 'slantDashDot' ? 18 * g : dashes ? dashes.reduce((sum, d) => sum + d, 0) * g : 4;
  const across = style === 'slantDashDot' ? 9 * g : 4;
  let pattern: CanvasPattern | null = null;
  if (doc && typeof ctx.createPattern === 'function') {
    const tile = doc.createElement('canvas');
    tile.width = horizontal ? along : across;
    tile.height = horizontal ? across : along;
    const tctx = tile.getContext('2d');
    if (tctx) {
      tctx.fillStyle = color;
      for (let a = 0; a < along; a += 1) {
        for (let c = 0; c < across; c += 1) {
          const ink = borderPixelInk(style, a, c, g) ?? 1;
          if (ink <= 0) continue;
          tctx.globalAlpha = ink;
          if (horizontal) tctx.fillRect(a, c, 1, 1);
          else tctx.fillRect(c, a, 1, 1);
        }
      }
      pattern = ctx.createPattern(tile, 'repeat');
    }
  }
  tiles.set(key, pattern);
  return pattern;
};

/** Single-pixel lines of a slant dash-dot stroke that ends a line. The
 *  spreadsheet closes the last dash on a slanted cut instead of a flat one:
 *  line `c` (counted from the stroke's first pixel across) ends where
 *  `along + 2c` reaches 1.5 gridline widths past the stroke's end, which trims
 *  the slow lines and extends the fast ones. */
export const slantTailLines = (rect: EdgeRect, horizontal: boolean, g: number): EdgeRect[] => {
  const reach = Math.round(1.5 * g);
  const lines: EdgeRect[] = [];
  const count = horizontal ? rect.h : rect.w;
  for (let c = 0; c < count; c += 1) {
    const length = (horizontal ? rect.w : rect.h) + reach - 2 * c;
    if (length <= 0) continue;
    lines.push(
      horizontal
        ? { x: rect.x, y: rect.y + c, w: length, h: 1 }
        : { x: rect.x + c, y: rect.y, w: 1, h: length },
    );
  }
  return lines;
};

/** Fill one stroke rectangle (device px, identity transform). Patterns are
 *  phased from `origin`, the grid's top-left corner in device px; slant
 *  strokes phase their across axis from `anchorRect` (default `rect`). */
export const fillEdgeRect = (
  ctx: CanvasRenderingContext2D,
  edge: BorderEdge,
  horizontal: boolean,
  rect: EdgeRect,
  color: string,
  g: number,
  origin: { x: number; y: number },
  anchorRect: EdgeRect = rect,
): void => {
  if (rect.w <= 0 || rect.h <= 0) return;
  if (borderPixelInk(edge.style, 0, 0, g) === null) {
    ctx.fillStyle = color;
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    return;
  }
  // Slant cuts restart on every stroke, so their across phase is the rect's.
  const slant = edge.style === 'slantDashDot';
  const anchor = {
    x: slant && !horizontal ? anchorRect.x : origin.x,
    y: slant && horizontal ? anchorRect.y : origin.y,
  };
  const pattern = patternTile(ctx, edge.style, horizontal, color, g);
  if (pattern) {
    ctx.save();
    ctx.translate(anchor.x, anchor.y);
    ctx.fillStyle = pattern;
    ctx.fillRect(rect.x - anchor.x, rect.y - anchor.y, rect.w, rect.h);
    ctx.restore();
    return;
  }
  ctx.fillStyle = color;
  for (let y = rect.y; y < rect.y + rect.h; y += 1) {
    for (let x = rect.x; x < rect.x + rect.w; x += 1) {
      const dx = x - anchor.x;
      const dy = y - anchor.y;
      const ink = horizontal
        ? borderPixelInk(edge.style, dx, dy, g)
        : borderPixelInk(edge.style, dy, dx, g);
      if (!ink) continue;
      ctx.globalAlpha = ink;
      ctx.fillRect(x, y, 1, 1);
    }
  }
  ctx.globalAlpha = 1;
};
