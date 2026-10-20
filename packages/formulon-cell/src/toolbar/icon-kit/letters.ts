/**
 * Letterforms used inside icons.
 *
 * Type is the one part of an icon that genuinely has to be drawn rather than
 * constructed, so these outlines are kept verbatim and only ever placed by
 * `letter()` — never re-typed at a new size. Counters are knocked out with the
 * even-odd fill rule so a letter stays legible on any surface.
 */

import { arc, ellipse, join, line, n, poly, vLine } from './path.js';
import { PALETTE, STROKE, type StrokeWeight } from './tokens.js';
import type { IconSegment } from './types.js';

export type Letterform = {
  /** Outline plus counters, in one path; painted with `evenodd`. */
  d: string;
  /** Natural bounding box as authored: [x, y, width, height]. */
  box: readonly [number, number, number, number];
};

export const LETTERS = {
  A: {
    d: join(
      'M5.5 16.5 10.4 4h3.2L18.5 16.5h-2.2l-1.1-3h-6.4l-1.1 3H5.5Z',
      'M9.5 11.5h5L12 4.6 9.5 11.5Z',
    ),
    box: [5.5, 4, 13, 12.5],
  },
  /** Semibold B: the stems are thinner than a display bold so it sits with I and U. */
  B: {
    d: join(
      'M0 0h4.6a3.6 3.1 0 0 1 0 6.2h.3a4.4 3.4 0 0 1 0 6.8H0Z',
      'M2.1 2.1h2.5a1.6 1.53 0 0 1 0 3.05H2.1Z',
      'M2.1 7.25h2.8a2.3 1.83 0 0 1 0 3.65H2.1Z',
    ),
    box: [0, 0, 9.3, 13],
  },
  I: {
    d: 'M10.1 5h8.1v1.2h-3.2l-4.3 10.6h3.9V18H6.5v-1.2h2.7l4.3-10.6h-3.4Z',
    box: [6.5, 5, 11.7, 13],
  },
  U: {
    d: 'M7.1 4.3h1.4v7.1c0 2.4 1.2 3.7 3.6 3.7s3.6-1.3 3.6-3.7V4.3h1.4v7.1c0 3.3-1.9 5.1-5 5.1s-5-1.8-5-5.1Z',
    box: [7.1, 4.3, 10, 12.2],
  },
  S: {
    d: 'M7.2 8.7c0-2 1.9-3.2 4.8-3.2 2.2 0 3.8.7 5 2l-1.7 1.7c-.8-.8-1.9-1.2-3.4-1.2-1.1 0-1.8.3-1.8.8 0 .6.8.9 2.9 1.4 2.8.7 4.4 1.8 4.4 4.1 0 2.4-2.2 4-5.3 4-2.6 0-4.5-.8-5.8-2.5l1.9-1.6c1 1.1 2.2 1.6 3.8 1.6 1.4 0 2.3-.5 2.3-1.3 0-.7-.7-1.1-2.7-1.6-2.8-.7-4.4-1.7-4.4-4.2Z',
    box: [6.3, 5.5, 11.1, 12.8],
  },
  Z: {
    d: 'M4.9 12.6h6.4v1.3l-4.1 4.2h4.3v1.4H4.8v-1.3l4.1-4.2H4.9v-1.4Z',
    box: [4.8, 12.6, 6.7, 6.9],
  },
  /** Summation sign, for the autosum family. */
  sigma: {
    d: 'M5.6 4.9h12.8v2.4H9.6l5.9 4.7-5.9 4.7h8.8v2.4H5.6v-2.2l7-4.9-7-4.9V4.9Z',
    box: [5.6, 4.9, 12.8, 14.2],
  },
  /** Lowercase-height A, for sort keys where two glyphs must stack. */
  aSmall: {
    d: join(
      'M4.7 11 7.3 4.5h1.6L11.5 11H9.8l-.5-1.4H6.9L6.4 11H4.7Z',
      'M7.3 8.3h1.6L8.1 6.1 7.3 8.3Z',
    ),
    box: [4.7, 4.5, 6.8, 6.5],
  },
  /** A literal number, for "find constants". */
  digits: {
    d: 'M8.75 9h1.6v7H8.85v-5.3l-1.1.7-.6-1.1L8.75 9ZM12.35 9h2.2c1.4 0 2.3.7 2.3 1.8 0 .7-.4 1.2-1 1.5.8.3 1.2.9 1.2 1.8 0 1.2-1 2-2.5 2h-2.2v-1.2h2.1c.7 0 1.1-.3 1.1-.9s-.4-.9-1.1-.9h-1.3v-1.1h1.2c.6 0 1-.3 1-.8s-.4-.8-1-.8h-2V9Z',
    box: [7.15, 9, 9.9, 7.1],
  },
} as const satisfies Record<string, Letterform>;

export type LetterName = keyof typeof LETTERS;

export type LetterOptions = {
  /** Cap height the letter should occupy. */
  height: number;
  /** Where the letter's own bounding box should be centred. */
  cx?: number;
  cy?: number;
  color?: string;
};

/** Place a letterform at a given size and centre. */
export const letter = (name: LetterName, options: LetterOptions): IconSegment => {
  const form = LETTERS[name];
  const [x, y, w, h] = form.box;
  const scale = options.height / h;
  const cx = options.cx ?? 12;
  const cy = options.cy ?? 12;
  const tx = cx - scale * (x + w / 2);
  const ty = cy - scale * (y + h / 2);
  return {
    d: form.d,
    fill: options.color ?? PALETTE.ink,
    fillRule: 'evenodd',
    transform: `translate(${round(tx)} ${round(ty)}) scale(${round(scale)})`,
  };
};

/** Width a letterform occupies once placed at `height`. */
export const letterWidth = (name: LetterName, height: number): number => {
  const [, , w, h] = LETTERS[name].box;
  return (w / h) * height;
};

const round = (value: number): number => Math.round(value * 1000) / 1000;

/**
 * Single-stroke letterforms, for the light type the ribbon draws in outline
 * (font size, alignment samples, sort keys, decimals).
 *
 * Each is a centreline drawn into a box rather than an outline placed by a
 * transform, so its stroke keeps a width from the token scale at any size.
 * `aspect` is the box's width over its height.
 */
type StrokeLetterform = {
  aspect: number;
  draw: (x: number, y: number, w: number, h: number) => string;
};

/** Height of the A crossbar, as a fraction of the cap height from the apex. */
const CROSSBAR = 0.64;

export const STROKE_LETTERS = {
  A: {
    aspect: 0.8,
    draw: (x, y, w, h) =>
      join(
        poly([
          [x, y + h],
          [x + w / 2, y],
          [x + w, y + h],
        ]),
        line(
          [x + (w / 2) * (1 - CROSSBAR), y + h * CROSSBAR],
          [x + (w / 2) * (1 + CROSSBAR), y + h * CROSSBAR],
        ),
      ),
  },
  /** Italic capital I with serifs. */
  I: {
    aspect: 0.56,
    draw: (x, y, w, h) =>
      join(
        line([x + w * 0.66, y], [x + w * 0.34, y + h]),
        line([x + w * 0.36, y], [x + w, y]),
        line([x, y + h], [x + w * 0.64, y + h]),
      ),
  },
  U: {
    aspect: 0.74,
    draw: (x, y, w, h) =>
      `${vLine(x, y, h - w / 2)}A${n(w / 2)} ${n(w / 2)} 0 0 0 ${n(x + w)} ${n(y + h - w / 2)}V${n(y)}`,
  },
  Z: {
    aspect: 0.74,
    draw: (x, y, w, h) =>
      poly([
        [x, y],
        [x + w, y],
        [x, y + h],
        [x + w, y + h],
      ]),
  },
  /** Summation sign, for the autosum family. */
  sigma: {
    aspect: 0.8,
    draw: (x, y, w, h) =>
      poly([
        [x + w, y],
        [x, y],
        [x + w * 0.56, y + h / 2],
        [x, y + h],
        [x + w, y + h],
      ]),
  },
  /** Lowercase a; `height` is the x-height. */
  a: {
    aspect: 0.84,
    draw: (x, y, w, h) =>
      join(ellipse(x + w * 0.42, y + h / 2, w * 0.42, h / 2), vLine(x + w, y, h)),
  },
  /** Lowercase b; `height` is the ascender height, the bowl takes the x-height. */
  b: {
    aspect: 0.56,
    draw: (x, y, w, h) => join(vLine(x, y, h), ellipse(x + w / 2, y + h * 0.7, w / 2, h * 0.3)),
  },
  /** Lowercase c; `height` is the x-height. */
  c: {
    aspect: 0.8,
    draw: (x, y, w, h) => arc(x + w / 2, y + h / 2, w / 2, h / 2, 42, 318),
  },
  zero: {
    aspect: 0.62,
    draw: (x, y, w, h) => ellipse(x + w / 2, y + h / 2, w / 2, h / 2),
  },
} as const satisfies Record<string, StrokeLetterform>;

export type StrokeLetterName = keyof typeof STROKE_LETTERS;

/** Lowercase ascender height relative to the x-height. */
export const ASCENDER = 1.45;

/** Draw a single-stroke letter at a given height and centre. */
export const strokeLetter = (
  name: StrokeLetterName,
  options: LetterOptions & { weight?: StrokeWeight },
): IconSegment => {
  const form: StrokeLetterform = STROKE_LETTERS[name];
  const h = options.height;
  const w = h * form.aspect;
  const cx = options.cx ?? 12;
  const cy = options.cy ?? 12;
  return {
    d: form.draw(cx - w / 2, cy - h / 2, w, h),
    fill: 'none',
    stroke: options.color ?? PALETTE.ink,
    strokeWidth: STROKE[options.weight ?? 'regular'],
    strokeLinecap: 'round',
    strokeLinejoin: 'round',
  };
};

/** Width a single-stroke letter's centreline occupies at `height`. */
export const strokeLetterWidth = (name: StrokeLetterName, height: number): number =>
  STROKE_LETTERS[name].aspect * height;
