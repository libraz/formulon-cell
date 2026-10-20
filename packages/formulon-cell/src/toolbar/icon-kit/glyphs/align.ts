/** Cell alignment, indent, wrap and merge icons. */

import { ASCENDER, strokeLetter, strokeLetterWidth } from '../letters.js';
import { join, rect, vLine } from '../path.js';
import { arrow, bars, headed, place, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/**
 * Text block used by the horizontal-alignment icons. The stack spans the full
 * optical box so the three icons balance against the rest of the set rather
 * than floating in the upper half.
 */
const TEXT = { x: 3, y: 4.4, width: 18, gap: 4.6 } as const;

/**
 * The vertical-alignment stack: long, short, long, centred, pushed to the edge
 * it aligns to. Position alone carries the meaning; no guide rule is drawn.
 */
const vStack = (y: number) => [
  ...bars({ x: 3, y, width: 18, widths: [1, 0.55, 1], gap: 4, align: 'center' }),
];

/** Lowercase x-height for the "ab" samples in wrap and orientation. */
const X_HEIGHT = 5.6;

const AB_GAP = 2;
const AB_HEIGHT = X_HEIGHT * ASCENDER;
const AB_WIDTH = strokeLetterWidth('a', X_HEIGHT) + AB_GAP + strokeLetterWidth('b', AB_HEIGHT);

/** "ab" set on a baseline, its left edge at `x`. */
const ab = (x: number, baseline: number, color: string = PALETTE.ink) => {
  const aw = strokeLetterWidth('a', X_HEIGHT);
  const bh = AB_HEIGHT;
  const bw = strokeLetterWidth('b', bh);
  return [
    strokeLetter('a', { height: X_HEIGHT, cx: x + aw / 2, cy: baseline - X_HEIGHT / 2, color }),
    strokeLetter('b', { height: bh, cx: x + aw + AB_GAP + bw / 2, cy: baseline - bh / 2, color }),
  ];
};

export const ALIGN_ICONS = {
  top: vStack(3.4),
  middle: vStack(8),
  bottomAlign: vStack(12.6),

  alignLeft: [...bars({ ...TEXT, widths: [1, 0.6, 1, 0.6], align: 'left' })],
  alignCenter: [...bars({ ...TEXT, widths: [1, 0.6, 1, 0.6], align: 'center' })],
  alignRight: [...bars({ ...TEXT, widths: [1, 0.6, 1, 0.6], align: 'right' })],

  indentIncrease: [
    ...bars({ x: 3, y: 3.6, width: 18, widths: [1], gap: 0 }),
    ...bars({ x: 10.4, y: 9, width: 10.6, widths: [1, 1], gap: 3.6 }),
    ...bars({ x: 3, y: 20.4, width: 18, widths: [1], gap: 0 }),
    ...arrow({ from: [3, 10.8], to: [8, 10.8], color: PALETTE.info, head: 'solid' }),
  ],
  indentDecrease: [
    ...bars({ x: 3, y: 3.6, width: 18, widths: [1], gap: 0 }),
    ...bars({ x: 10.4, y: 9, width: 10.6, widths: [1, 1], gap: 3.6 }),
    ...bars({ x: 3, y: 20.4, width: 18, widths: [1], gap: 0 }),
    ...arrow({ from: [8, 10.8], to: [3, 10.8], color: PALETTE.info, head: 'solid' }),
  ],

  // "ab" runs to the edge and the return carries the rest onto the next line.
  wrap: [
    ...ab(3.4, 10.4),
    strokeLetter('c', {
      height: X_HEIGHT,
      cx: 3.4 + strokeLetterWidth('c', X_HEIGHT) / 2,
      cy: 20 - X_HEIGHT / 2,
    }),
    ...headed({
      shaft: 'M15.4 8.2h1.8a4.2 4.2 0 0 1 0 8.4h-5',
      tip: [10.6, 16.6],
      headDir: [-1, 0],
      color: PALETTE.info,
      weight: 'regular',
    }),
  ],

  // The text and the arrow turned together: the angle is the message.
  textOrientation: [
    ...place(ab(0, AB_HEIGHT), {
      box: [0, 0, AB_WIDTH, AB_HEIGHT],
      size: AB_WIDTH,
      cx: 8,
      cy: 8,
      rotate: -45,
    }),
    ...arrow({ from: [4.8, 20.4], to: [20.4, 4.8], color: PALETTE.info, weight: 'regular' }),
  ],

  // Two cells become one: the divider stops short and the content spreads.
  merge: [
    stroked(rect(2.4, 5.4, 19.2, 13.2), PALETTE.ink, 'regular', { cap: 'butt', join: 'miter' }),
    stroked(join(vLine(12, 5.4, 3), vLine(12, 15.6, 3)), PALETTE.ink, 'thin', { cap: 'butt' }),
    ...arrow({
      from: [5.2, 12],
      to: [18.8, 12],
      color: PALETTE.info,
      head: 'chevron',
      tail: 'chevron',
      weight: 'regular',
    }),
  ],
} satisfies Record<string, IconDefinition>;
