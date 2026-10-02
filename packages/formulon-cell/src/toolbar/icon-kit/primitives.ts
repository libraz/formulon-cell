/**
 * Composable icon parts.
 *
 * Every ribbon glyph is assembled from these rather than from literal path
 * data. The point is that shared structure — an arrow head, a sheet frame, a
 * corner badge — is computed from one definition, so it cannot drift between
 * icons the way copy-pasted coordinates did.
 */

import {
  advance,
  circle,
  direction,
  ellipse,
  hRules,
  join,
  line,
  n,
  type Point,
  perpendicular,
  poly,
  rect,
  roundRect,
  vRules,
} from './path.js';
import { CENTER, PALETTE, STROKE, type StrokeWeight, strokeWidth } from './tokens.js';
import type { IconSegment } from './types.js';

/** A filled path. */
export const filled = (d: string, fill: string): IconSegment => ({ d, fill });

/** A stroked path with the round caps and joins the set uses throughout. */
export const stroked = (
  d: string,
  stroke: string,
  weight: StrokeWeight = 'regular',
  options: { cap?: 'butt' | 'round' | 'square'; join?: 'round' | 'miter'; dash?: string } = {},
): IconSegment => ({
  d,
  fill: 'none',
  stroke,
  strokeWidth: STROKE[weight],
  strokeLinecap: options.cap ?? 'round',
  strokeLinejoin: options.join ?? 'round',
  ...(options.dash ? { strokeDasharray: options.dash } : {}),
});

/** A filled path with an outline, for solid shapes that need definition. */
export const outlined = (
  d: string,
  fill: string,
  stroke: string,
  weight: StrokeWeight = 'regular',
): IconSegment => ({
  d,
  fill,
  stroke,
  strokeWidth: STROKE[weight],
  strokeLinejoin: 'round',
});

/**
 * Rescale and reposition a group of segments.
 *
 * `box` is the group's natural bounding box; the group is scaled so its longest
 * axis matches `size` and its centre lands on (`cx`, `cy`). Stroke widths scale
 * with it, which is what keeps a shrunk-down mark from looking heavy.
 */
export const place = (
  segments: readonly IconSegment[],
  options: {
    box: readonly [number, number, number, number];
    size: number;
    cx?: number;
    cy?: number;
    /** Clockwise rotation in degrees, about the group's own centre. */
    rotate?: number;
  },
): IconSegment[] => {
  const [x, y, w, h] = options.box;
  const scale = options.size / Math.max(w, h);
  const cx = options.cx ?? CENTER;
  const cy = options.cy ?? CENTER;
  const spin = options.rotate ? ` rotate(${n(options.rotate)})` : '';
  const transform =
    `translate(${n(cx)} ${n(cy)})${spin} scale(${n(scale)})` +
    ` translate(${n(-(x + w / 2))} ${n(-(y + h / 2))})`;
  return segments.map((segment) => ({
    ...segment,
    transform: segment.transform ? `${transform} ${segment.transform}` : transform,
  }));
};

/** A box of the given size centred on the canvas. */
export const centeredBox = (
  w: number,
  h: number = w,
): { x: number; y: number; w: number; h: number } => ({
  x: CENTER - w / 2,
  y: CENTER - h / 2,
  w,
  h,
});

// ---------------------------------------------------------------------------
// Arrows
// ---------------------------------------------------------------------------

export type ArrowHead = 'chevron' | 'solid' | 'none';

export type ArrowOptions = {
  from: Point;
  to: Point;
  color: string;
  weight?: StrokeWeight;
  /** Head at the `to` end. */
  head?: ArrowHead;
  /** Head at the `from` end, for double-headed arrows. */
  tail?: ArrowHead;
  /** Scales the canonical head. Only for arrows that must carry the icon. */
  headScale?: number;
};

/** Head length and half-width as multiples of the shaft width. */
const HEAD_LENGTH_RATIO = 2.1;
const HEAD_HALF_RATIO = 1.5;

const headPoints = (
  tip: Point,
  dir: Point,
  width: number,
  scale: number,
): { base: Point; left: Point; right: Point } => {
  const length = width * HEAD_LENGTH_RATIO * scale;
  const half = width * HEAD_HALF_RATIO * scale;
  const perp = perpendicular(dir);
  const base = advance(tip, dir, -length);
  return {
    base,
    left: [base[0] + perp[0] * half, base[1] + perp[1] * half],
    right: [base[0] - perp[0] * half, base[1] - perp[1] * half],
  };
};

/**
 * A straight arrow.
 *
 * The head apex always sits exactly on the shaft end: a chevron head shares the
 * shaft's end point, and a solid head has the shaft stop short of its base.
 * Hand-drawn heads used to overshoot or float away from the shaft, which is the
 * single most visible defect in the old set.
 */
export const arrow = (options: ArrowOptions): IconSegment[] => {
  const { from, to, color } = options;
  const weight = options.weight ?? 'bold';
  const width = strokeWidth(weight);
  const scale = options.headScale ?? 1;
  const head = options.head ?? 'chevron';
  const tail = options.tail ?? 'none';

  const forward = direction(from, to);
  const backward: Point = [-forward[0], -forward[1]];

  const tipGeom = head === 'none' ? null : headPoints(to, forward, width, scale);
  const tailGeom = tail === 'none' ? null : headPoints(from, backward, width, scale);

  // A solid head hides the shaft end inside the triangle; a chevron shares it.
  const shaftEnd = head === 'solid' && tipGeom ? advance(to, forward, -width * scale) : to;
  const shaftStart = tail === 'solid' && tailGeom ? advance(from, backward, -width * scale) : from;

  const segments: IconSegment[] = [stroked(line(shaftStart, shaftEnd), color, weight)];

  for (const [end, geom, kind] of [
    [to, tipGeom, head],
    [from, tailGeom, tail],
  ] as const) {
    if (!geom || kind === 'none') continue;
    if (kind === 'solid') {
      segments.push(filled(poly([end, geom.left, geom.right], true), color));
    } else {
      segments.push(stroked(poly([geom.left, end, geom.right]), color, weight));
    }
  }

  return segments;
};

/**
 * An arrow along an arbitrary shaft path.
 *
 * The shaft is given as raw path data — an L-bend, an arc, a hook — and the
 * head is placed at `tip` pointing along `headDir`. Use this whenever the shaft
 * is not a straight line; it keeps every head in the set built the same way.
 */
export const headed = (options: {
  shaft: string;
  tip: Point;
  headDir: Point;
  color: string;
  weight?: StrokeWeight;
  head?: Exclude<ArrowHead, 'none'>;
  headScale?: number;
}): IconSegment[] => {
  const weight = options.weight ?? 'bold';
  const width = strokeWidth(weight);
  const scale = options.headScale ?? 1;
  const geom = headPoints(options.tip, options.headDir, width, scale);
  const kind = options.head ?? 'solid';
  return [
    stroked(options.shaft, options.color, weight),
    kind === 'solid'
      ? filled(poly([options.tip, geom.left, geom.right], true), options.color)
      : stroked(poly([geom.left, options.tip, geom.right]), options.color, weight),
  ];
};

/**
 * A quarter-turn arrow following a circular arc, for undo/redo and refresh.
 * `sweep` follows the SVG arc flag: 1 turns clockwise.
 */
export const arcArrow = (options: {
  from: Point;
  to: Point;
  radius: number;
  sweep: 0 | 1;
  /** Direction the head points, as a unit vector at `to`. */
  headDir: Point;
  color: string;
  weight?: StrokeWeight;
  head?: ArrowHead;
  headScale?: number;
  largeArc?: 0 | 1;
}): IconSegment[] => {
  const weight = options.weight ?? 'bold';
  const width = strokeWidth(weight);
  const scale = options.headScale ?? 1;
  const head = options.head ?? 'solid';
  const geom = headPoints(options.to, options.headDir, width, scale);
  const end = head === 'solid' ? geom.base : options.to;
  const r = n(options.radius);
  const arc = `M${n(options.from[0])} ${n(options.from[1])}A${r} ${r} 0 ${options.largeArc ?? 0} ${options.sweep} ${n(end[0])} ${n(end[1])}`;

  const segments: IconSegment[] = [stroked(arc, options.color, weight)];
  if (head === 'solid') {
    segments.push(filled(poly([options.to, geom.left, geom.right], true), options.color));
  } else if (head === 'chevron') {
    segments.push(stroked(poly([geom.left, options.to, geom.right]), options.color, weight));
  }
  return segments;
};

// ---------------------------------------------------------------------------
// Corner badge
// ---------------------------------------------------------------------------

export type BadgeCorner = 'tr' | 'br' | 'tl' | 'bl';
export type BadgeGlyph = 'plus' | 'minus' | 'cross' | 'check' | 'star' | 'pencil' | 'dots' | 'bang';

export type BadgeOptions = {
  glyph: BadgeGlyph;
  corner?: BadgeCorner;
  /** Badge surface. Defaults to the accent colour for additive glyphs. */
  tone?: string;
  /** Glyph colour on top of the surface. */
  onTone?: string;
  shape?: 'square' | 'circle';
};

/** Badge plate size and its inset from the canvas edge. */
const BADGE_SIZE = 8;
const BADGE_INSET = 1;
const BADGE_GLYPH_SCALE = BADGE_SIZE / 9;

const DEFAULT_BADGE_TONE: Record<BadgeGlyph, string> = {
  plus: PALETTE.accent,
  minus: PALETTE.danger,
  cross: PALETTE.danger,
  check: PALETTE.accent,
  star: PALETTE.warn,
  pencil: PALETTE.info,
  dots: PALETTE.grid,
  bang: PALETTE.warn,
};

const badgeOrigin = (corner: BadgeCorner): Point => {
  const far = 24 - BADGE_INSET - BADGE_SIZE;
  switch (corner) {
    case 'tr':
      return [far, BADGE_INSET];
    case 'br':
      return [far, far];
    case 'tl':
      return [BADGE_INSET, BADGE_INSET];
    case 'bl':
      return [BADGE_INSET, far];
  }
};

/**
 * A status badge pinned to a canvas corner.
 *
 * The base glyph keeps its own centre and simply shrinks to `OPTICAL_BADGED`;
 * it is never shifted away from the badge. Nudging the base is what made the
 * old insert/delete family look off-balance.
 */
export const badge = (options: BadgeOptions): IconSegment[] => {
  const corner = options.corner ?? 'tr';
  const [x, y] = badgeOrigin(corner);
  const tone = options.tone ?? DEFAULT_BADGE_TONE[options.glyph];
  const onTone = options.onTone ?? PALETTE.paper;
  const cx = x + BADGE_SIZE / 2;
  const cy = y + BADGE_SIZE / 2;
  const plate =
    options.shape === 'circle'
      ? filled(circle(cx, cy, BADGE_SIZE / 2), tone)
      : filled(roundRect(x, y, BADGE_SIZE, BADGE_SIZE, 1.6), tone);

  const arm = 2.1 * BADGE_GLYPH_SCALE;
  const glyph = ((): IconSegment[] => {
    switch (options.glyph) {
      case 'plus':
        return [
          stroked(
            join(
              `M${n(cx - arm)} ${n(cy)}h${n(arm * 2)}`,
              `M${n(cx)} ${n(cy - arm)}v${n(arm * 2)}`,
            ),
            onTone,
            'heavy',
          ),
        ];
      case 'minus':
        return [stroked(`M${n(cx - arm)} ${n(cy)}h${n(arm * 2)}`, onTone, 'heavy')];
      case 'cross':
        return [
          stroked(
            join(
              line([cx - arm * 0.8, cy - arm * 0.8], [cx + arm * 0.8, cy + arm * 0.8]),
              line([cx + arm * 0.8, cy - arm * 0.8], [cx - arm * 0.8, cy + arm * 0.8]),
            ),
            onTone,
            'heavy',
          ),
        ];
      case 'check':
        return [
          stroked(
            poly([
              [cx - 2 * BADGE_GLYPH_SCALE, cy],
              [cx - 0.6 * BADGE_GLYPH_SCALE, cy + 1.5 * BADGE_GLYPH_SCALE],
              [cx + 2.1 * BADGE_GLYPH_SCALE, cy - 1.7 * BADGE_GLYPH_SCALE],
            ]),
            onTone,
            'heavy',
          ),
        ];
      case 'star':
        return [filled(star(cx, cy, 3.1 * BADGE_GLYPH_SCALE, 1.4 * BADGE_GLYPH_SCALE), onTone)];
      case 'pencil': {
        const scale = BADGE_GLYPH_SCALE;
        const tip: Point = [cx - 3.25 * scale, cy + 3.25 * scale];
        const nibBase: Point = [cx - 2.25 * scale, cy + 2.25 * scale];
        const bodyEnd: Point = [cx + 2.75 * scale, cy - 2.75 * scale];
        const capEnd: Point = [cx + 3.75 * scale, cy - 3.75 * scale];
        const side: Point = [0.7 * scale, 0.7 * scale];
        const silhouette = poly(
          [
            tip,
            [nibBase[0] + side[0], nibBase[1] + side[1]],
            [capEnd[0] + side[0], capEnd[1] + side[1]],
            [capEnd[0] - side[0], capEnd[1] - side[1]],
            [nibBase[0] - side[0], nibBase[1] - side[1]],
          ],
          true,
        );
        return [
          filled(silhouette, onTone),
          stroked(
            line(
              [bodyEnd[0] + side[0], bodyEnd[1] + side[1]],
              [bodyEnd[0] - side[0], bodyEnd[1] - side[1]],
            ),
            tone,
            'hairline',
            { cap: 'butt' },
          ),
        ];
      }
      case 'bang':
        return [
          stroked(
            join(
              `M${n(cx)} ${n(cy - 2.4 * BADGE_GLYPH_SCALE)}v${n(2.6 * BADGE_GLYPH_SCALE)}`,
              `M${n(cx)} ${n(cy + 2.2 * BADGE_GLYPH_SCALE)}v${n(0.1 * BADGE_GLYPH_SCALE)}`,
            ),
            onTone,
            'heavy',
          ),
        ];
      case 'dots':
        return [
          filled(
            join(
              circle(cx - 2.3 * BADGE_GLYPH_SCALE, cy, 0.75 * BADGE_GLYPH_SCALE),
              circle(cx, cy, 0.75 * BADGE_GLYPH_SCALE),
              circle(cx + 2.3 * BADGE_GLYPH_SCALE, cy, 0.75 * BADGE_GLYPH_SCALE),
            ),
            onTone,
          ),
        ];
    }
  })();

  return [plate, ...glyph];
};

/** Five-pointed star path. */
export const star = (cx: number, cy: number, outer: number, inner: number): string => {
  const points: Point[] = [];
  for (let i = 0; i < 10; i += 1) {
    const r = i % 2 === 0 ? outer : inner;
    const angle = (Math.PI / 5) * i - Math.PI / 2;
    points.push([cx + Math.cos(angle) * r, cy + Math.sin(angle) * r]);
  }
  return poly(points, true);
};

// ---------------------------------------------------------------------------
// Sheet / table
// ---------------------------------------------------------------------------

export type SheetBands = readonly SheetBand[];

export type SheetBand = {
  axis: 'row' | 'col';
  /** Zero-based index, or a [start, end] span. */
  index: number | readonly [number, number];
  fill: string;
};

export type SheetOptions = {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  cols?: number;
  rows?: number;
  frame?: string;
  frameWeight?: StrokeWeight;
  surface?: string;
  rules?: string;
  bands?: readonly SheetBand[];
};

/**
 * A spreadsheet grid: surface, optional highlighted bands, interior rules, then
 * the frame on top so the outline is never broken by a band edge.
 */
export const sheet = (options: SheetOptions = {}): IconSegment[] => {
  const x = options.x ?? 2;
  const y = options.y ?? 2;
  const w = options.w ?? 20;
  const h = options.h ?? 20;
  const cols = options.cols ?? 3;
  const rows = options.rows ?? 3;
  const colWidth = w / cols;
  const rowHeight = h / rows;

  const segments: IconSegment[] = [filled(rect(x, y, w, h), options.surface ?? PALETTE.paper)];

  for (const band of options.bands ?? []) {
    const [start, end] = typeof band.index === 'number' ? [band.index, band.index] : band.index;
    const span = end - start + 1;
    segments.push(
      band.axis === 'row'
        ? filled(rect(x, y + rowHeight * start, w, rowHeight * span), band.fill)
        : filled(rect(x + colWidth * start, y, colWidth * span, h), band.fill),
    );
  }

  if (rows > 1 || cols > 1) {
    segments.push(
      stroked(
        join(
          rows > 1 ? hRules(x, y + rowHeight, w, rowHeight, rows - 1) : '',
          cols > 1 ? vRules(x + colWidth, y, h, colWidth, cols - 1) : '',
        ),
        options.rules ?? PALETTE.gridLight,
        'hairline',
        { cap: 'butt' },
      ),
    );
  }

  segments.push(
    stroked(rect(x, y, w, h), options.frame ?? PALETTE.ink, options.frameWeight ?? 'regular', {
      cap: 'butt',
      join: 'miter',
    }),
  );

  return segments;
};

// ---------------------------------------------------------------------------
// Document, clipboard, screen
// ---------------------------------------------------------------------------

export type DocOptions = {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  /** Size of the folded top-right corner; 0 draws a plain rectangle. */
  fold?: number;
  surface?: string;
  frame?: string;
  lines?: number;
  lineColor?: string;
};

/** A page with an optional folded corner and text rules. */
export const doc = (options: DocOptions = {}): IconSegment[] => {
  const x = options.x ?? 4;
  const y = options.y ?? 2;
  const w = options.w ?? 16;
  const h = options.h ?? 20;
  const fold = options.fold ?? 4;
  const surface = options.surface ?? PALETTE.paper;
  const frame = options.frame ?? PALETTE.ink;

  const body = fold
    ? poly(
        [
          [x, y],
          [x + w - fold, y],
          [x + w, y + fold],
          [x + w, y + h],
          [x, y + h],
        ],
        true,
      )
    : rect(x, y, w, h);

  const segments: IconSegment[] = [outlined(body, surface, frame, 'regular')];
  if (fold) {
    segments.push(
      stroked(
        poly([
          [x + w - fold, y],
          [x + w - fold, y + fold],
          [x + w, y + fold],
        ]),
        frame,
        'thin',
        { join: 'miter' },
      ),
    );
  }

  const lines = options.lines ?? 0;
  if (lines > 0) {
    const top = y + (fold ? fold + 2.6 : 4);
    const gap = 2.8;
    segments.push(
      ...textLines({
        x: x + 2.6,
        y: top,
        width: w - 5.2,
        count: lines,
        gap,
        color: options.lineColor ?? PALETTE.grid,
      }),
    );
  }
  return segments;
};

/** A clipboard: board, clip plate, and clip tab. */
export const clipboard = (
  options: { surface?: string; frame?: string; clip?: string } = {},
): IconSegment[] => {
  const surface = options.surface ?? PALETTE.paper;
  const frame = options.frame ?? PALETTE.ink;
  const clip = options.clip ?? PALETTE.tan;
  return [
    outlined(roundRect(4, 4, 16, 18, 1.6), surface, frame, 'regular'),
    outlined(roundRect(7.5, 1.8, 9, 4.4, 1.2), clip, PALETTE.tanDeep, 'thin'),
    outlined(roundRect(9.75, 1, 4.5, 2.4, 0.8), PALETTE.warnSoft, PALETTE.tanDeep, 'hairline'),
  ];
};

/** A display: screen panel and stand. */
export const monitor = (
  options: { surface?: string; frame?: string; stand?: boolean } = {},
): IconSegment[] => {
  const segments: IconSegment[] = [
    outlined(
      roundRect(2, 4, 20, 14, 1.4),
      options.surface ?? PALETTE.paper,
      options.frame ?? PALETTE.ink,
      'regular',
    ),
  ];
  if (options.stand !== false) {
    segments.push(stroked(join(`M12 18v2.4`, `M8.4 21h7.2`), options.frame ?? PALETTE.ink, 'bold'));
  }
  return segments;
};

/** Evenly spaced text rules. */
export const textLines = (options: {
  x: number;
  y: number;
  width: number;
  count: number;
  gap?: number;
  color?: string;
  weight?: StrokeWeight;
  /** Fraction of `width` used by the last rule, so text reads as ragged. */
  lastRatio?: number;
}): IconSegment[] => {
  const gap = options.gap ?? 3;
  const last = options.lastRatio ?? 0.65;
  const parts = Array.from({ length: options.count }, (_, i) => {
    const width =
      i === options.count - 1 && options.count > 1 ? options.width * last : options.width;
    return `M${n(options.x)} ${n(options.y + gap * i)}h${n(width)}`;
  });
  return [stroked(join(...parts), options.color ?? PALETTE.grid, options.weight ?? 'thin')];
};

/**
 * A stack of rules standing in for lines of text, each a fraction of the full
 * measure so the block reads as prose rather than as a table.
 */
export const bars = (options: {
  x: number;
  y: number;
  width: number;
  widths: readonly number[];
  gap?: number;
  align?: 'left' | 'center' | 'right';
  color?: string;
  weight?: StrokeWeight;
}): IconSegment[] => {
  const gap = options.gap ?? 3.4;
  const align = options.align ?? 'left';
  const parts = options.widths.map((fraction, i) => {
    const w = options.width * fraction;
    const slack = options.width - w;
    const offset = align === 'left' ? 0 : align === 'right' ? slack : slack / 2;
    return `M${n(options.x + offset)} ${n(options.y + gap * i)}h${n(w)}`;
  });
  return [stroked(join(...parts), options.color ?? PALETTE.ink, options.weight ?? 'bold')];
};

// ---------------------------------------------------------------------------
// Recurring symbols
// ---------------------------------------------------------------------------

/** The filter cone, sized to the optical box. */
export const funnel = (
  options: { surface?: string; frame?: string; scale?: number } = {},
): IconSegment[] => {
  const s = options.scale ?? 1;
  const cx = CENTER;
  const halfTop = 9 * s;
  const top = CENTER - 9 * s;
  const neckHalf = 1.6 * s;
  const neckTop = top + 8.4 * s;
  const bottom = CENTER + 9 * s;
  const d = poly(
    [
      [cx - halfTop, top],
      [cx + halfTop, top],
      [cx + neckHalf, neckTop],
      [cx + neckHalf, bottom - 1.6 * s],
      [cx - neckHalf, bottom],
      [cx - neckHalf, neckTop],
    ],
    true,
  );
  return [outlined(d, options.surface ?? PALETTE.paper, options.frame ?? PALETTE.ink, 'regular')];
};

/** A magnifier: lens and handle, drawn so the handle meets the rim exactly. */
export const magnifier = (
  options: {
    cx?: number;
    cy?: number;
    r?: number;
    handle?: number;
    surface?: string;
    frame?: string;
    handleColor?: string;
  } = {},
): IconSegment[] => {
  const cx = options.cx ?? 10.4;
  const cy = options.cy ?? 10.4;
  const r = options.r ?? 6;
  const handle = options.handle ?? 5.4;
  const unit = Math.SQRT1_2;
  const start: Point = [cx + unit * r, cy + unit * r];
  const end: Point = [start[0] + unit * handle, start[1] + unit * handle];
  return [
    outlined(
      circle(cx, cy, r),
      options.surface ?? PALETTE.paper,
      options.frame ?? PALETTE.ink,
      'regular',
    ),
    stroked(line(start, end), options.handleColor ?? options.frame ?? PALETTE.ink, 'heavy'),
  ];
};

/** A chain link, angled along the icon's diagonal. */
export const chainLink = (options: {
  from: Point;
  to: Point;
  color?: string;
  radius?: number;
}): IconSegment => {
  const r = options.radius ?? 2.6;
  const dir = direction(options.from, options.to);
  const perp = perpendicular(dir);
  const a = options.from;
  const b = options.to;
  // Both caps sweep away from the link's own axis; sweeping the other way
  // folds them inward and the capsule collapses into a bowtie.
  const d = join(
    `M${n(a[0] + perp[0] * r)} ${n(a[1] + perp[1] * r)}`,
    `L${n(b[0] + perp[0] * r)} ${n(b[1] + perp[1] * r)}`,
    `A${n(r)} ${n(r)} 0 0 0 ${n(b[0] - perp[0] * r)} ${n(b[1] - perp[1] * r)}`,
    `L${n(a[0] - perp[0] * r)} ${n(a[1] - perp[1] * r)}`,
    `A${n(r)} ${n(r)} 0 0 0 ${n(a[0] + perp[0] * r)} ${n(a[1] + perp[1] * r)}`,
    'z',
  );
  return stroked(d, options.color ?? PALETTE.info, 'bold');
};

/** An ellipse used as an annotation ring. */
export const ring = (
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  color: string,
  weight: StrokeWeight = 'bold',
): IconSegment => stroked(ellipse(cx, cy, rx, ry), color, weight);
