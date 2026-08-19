/** Cell alignment, indent, wrap and merge icons. */

import { letter } from '../letters.js';
import { hLine } from '../path.js';
import { arrow, bars, headed, place, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/**
 * Text block used by the horizontal-alignment icons. The stack spans the full
 * optical box so the three icons balance against the rest of the set rather
 * than floating in the upper half.
 */
const TEXT = { x: 3, y: 4.4, width: 18, gap: 4.6 } as const;

/** Rule marking the edge the content is pushed against. */
const guide = (d: string) => stroked(d, PALETTE.accent, 'heavy');

export const ALIGN_ICONS = {
  top: [
    guide(hLine(2.4, 3.2, 19.2)),
    ...bars({ x: 4.4, y: 8.8, width: 15.2, widths: [1, 0.7, 0.85], gap: 5.2, color: PALETTE.ink }),
  ],
  middle: [
    ...bars({ x: 4.4, y: 3.6, width: 15.2, widths: [0.85, 1], gap: 4.2, color: PALETTE.ink }),
    guide(hLine(2.4, 12, 19.2)),
    ...bars({ x: 4.4, y: 16.2, width: 15.2, widths: [1, 0.7], gap: 4.2, color: PALETTE.ink }),
  ],
  bottomAlign: [
    ...bars({ x: 4.4, y: 4.8, width: 15.2, widths: [1, 0.7, 0.85], gap: 5.2, color: PALETTE.ink }),
    guide(hLine(2.4, 20.8, 19.2)),
  ],

  alignLeft: [...bars({ ...TEXT, widths: [1, 0.6, 0.95, 0.55], align: 'left' })],
  alignCenter: [...bars({ ...TEXT, widths: [1, 0.6, 0.95, 0.55], align: 'center' })],
  alignRight: [...bars({ ...TEXT, widths: [1, 0.6, 0.95, 0.55], align: 'right' })],

  indentIncrease: [
    ...bars({ x: 3, y: 3.6, width: 18, widths: [1], gap: 0 }),
    ...bars({ x: 10.4, y: 9, width: 10.6, widths: [1, 1], gap: 3.6 }),
    ...bars({ x: 3, y: 20.4, width: 18, widths: [1], gap: 0 }),
    ...arrow({ from: [3, 14.4], to: [8, 14.4], color: PALETTE.accent, head: 'solid' }),
  ],
  indentDecrease: [
    ...bars({ x: 3, y: 3.6, width: 18, widths: [1], gap: 0 }),
    ...bars({ x: 10.4, y: 9, width: 10.6, widths: [1, 1], gap: 3.6 }),
    ...bars({ x: 3, y: 20.4, width: 18, widths: [1], gap: 0 }),
    ...arrow({ from: [8, 14.4], to: [3, 14.4], color: PALETTE.accent, head: 'solid' }),
  ],

  // Text runs to the edge and turns back onto the next line.
  wrap: [
    ...bars({ x: 3, y: 5.4, width: 18, widths: [1], gap: 0 }),
    ...headed({
      shaft: 'M3 12.4h13.6a3.4 3.4 0 0 1 0 6.8h-6.2',
      tip: [8.4, 19.2],
      headDir: [-1, 0],
      color: PALETTE.accent,
      weight: 'regular',
    }),
  ],

  // The letter and its baseline turned together: the angle is the message.
  textOrientation: [
    ...place(
      [
        letter('A', { height: 12.6, cx: 12, cy: 9.4 }),
        stroked('M2.6 18h18.8', PALETTE.accent, 'heavy'),
      ],
      { box: [2.6, 3.1, 18.8, 15.4], size: 17.2, rotate: -38 },
    ),
  ],

  merge: [
    ...sheet({ x: 2.4, y: 5.4, w: 19.2, h: 13.2, cols: 2, rows: 1, rules: PALETTE.gridLight }),
    ...arrow({
      from: [11.4, 12],
      to: [5.4, 12],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
    ...arrow({
      from: [12.6, 12],
      to: [18.6, 12],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
  ],
} satisfies Record<string, IconDefinition>;
