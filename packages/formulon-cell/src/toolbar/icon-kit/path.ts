/**
 * Path-data builders.
 *
 * These emit `d` strings only — no colour, no stroke. Everything above this
 * layer composes shapes from these instead of typing coordinates by hand, so a
 * rectangle is always closed the same way and an arc always winds the same
 * direction.
 */

export type Point = readonly [number, number];

/** Format a coordinate: two decimals max, no negative zero, no trailing zeros. */
export const n = (value: number): string => {
  const rounded = Math.round(value * 100) / 100;
  return String(Object.is(rounded, -0) ? 0 : rounded);
};

/** Axis-aligned rectangle. */
export const rect = (x: number, y: number, w: number, h: number): string =>
  `M${n(x)} ${n(y)}h${n(w)}v${n(h)}h${n(-w)}z`;

/** Rectangle with uniform corner radius. Falls back to `rect` when r <= 0. */
export const roundRect = (x: number, y: number, w: number, h: number, r: number): string => {
  const radius = Math.min(r, w / 2, h / 2);
  if (radius <= 0) return rect(x, y, w, h);
  const a = `a${n(radius)} ${n(radius)} 0 0 1`;
  return [
    `M${n(x + radius)} ${n(y)}`,
    `h${n(w - radius * 2)}`,
    `${a} ${n(radius)} ${n(radius)}`,
    `v${n(h - radius * 2)}`,
    `${a} ${n(-radius)} ${n(radius)}`,
    `h${n(-(w - radius * 2))}`,
    `${a} ${n(-radius)} ${n(-radius)}`,
    `v${n(-(h - radius * 2))}`,
    `${a} ${n(radius)} ${n(-radius)}`,
    'z',
  ].join('');
};

/** Full circle, wound clockwise. */
export const circle = (cx: number, cy: number, r: number): string =>
  `M${n(cx - r)} ${n(cy)}a${n(r)} ${n(r)} 0 1 0 ${n(r * 2)} 0a${n(r)} ${n(r)} 0 1 0 ${n(-r * 2)} 0z`;

/** Full ellipse, wound clockwise. */
export const ellipse = (cx: number, cy: number, rx: number, ry: number): string =>
  `M${n(cx - rx)} ${n(cy)}a${n(rx)} ${n(ry)} 0 1 0 ${n(rx * 2)} 0a${n(rx)} ${n(ry)} 0 1 0 ${n(-rx * 2)} 0z`;

/** Polyline through `points`, optionally closed. */
export const poly = (points: readonly Point[], close = false): string => {
  const head = points[0];
  if (!head) return '';
  const tail = points.slice(1);
  const body = tail.map(([x, y]) => `L${n(x)} ${n(y)}`).join('');
  return `M${n(head[0])} ${n(head[1])}${body}${close ? 'z' : ''}`;
};

/** Straight segment between two points. */
export const line = (from: Point, to: Point): string => poly([from, to]);

/** Horizontal rule. */
export const hLine = (x: number, y: number, length: number): string =>
  `M${n(x)} ${n(y)}h${n(length)}`;

/** Vertical rule. */
export const vLine = (x: number, y: number, length: number): string =>
  `M${n(x)} ${n(y)}v${n(length)}`;

/** Evenly spaced horizontal rules, top-down. */
export const hRules = (x: number, y: number, length: number, gap: number, count: number): string =>
  Array.from({ length: count }, (_, i) => hLine(x, y + gap * i, length)).join('');

/** Evenly spaced vertical rules, left to right. */
export const vRules = (x: number, y: number, length: number, gap: number, count: number): string =>
  Array.from({ length: count }, (_, i) => vLine(x + gap * i, y, length)).join('');

/** Join several path fragments into one `d`. */
export const join = (...parts: readonly (string | null | undefined | false)[]): string =>
  parts.filter((part): part is string => Boolean(part)).join('');

/** Unit vector from `from` towards `to`; returns [1, 0] for a zero-length pair. */
export const direction = (from: Point, to: Point): Point => {
  const dx = to[0] - from[0];
  const dy = to[1] - from[1];
  const length = Math.hypot(dx, dy);
  return length === 0 ? [1, 0] : [dx / length, dy / length];
};

/** Move a point along a direction by `distance`. */
export const advance = (point: Point, dir: Point, distance: number): Point => [
  point[0] + dir[0] * distance,
  point[1] + dir[1] * distance,
];

/** Rotate a direction 90 degrees clockwise. */
export const perpendicular = (dir: Point): Point => [-dir[1], dir[0]];
