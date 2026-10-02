/** Review, accessibility, protection and annotation icons. */

import { letter } from '../letters.js';
import { circle, join, line, poly, rect, roundRect } from '../path.js';
import {
  badge,
  bars,
  chainLink,
  doc,
  filled,
  magnifier,
  outlined,
  stroked,
} from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** The comment bubble, with its tail on the lower left. */
const bubble = (box: { x: number; y: number; w: number; h: number }) =>
  poly(
    [
      [box.x, box.y],
      [box.x + box.w, box.y],
      [box.x + box.w, box.y + box.h],
      [box.x + box.w * 0.42, box.y + box.h],
      [box.x + box.w * 0.2, box.y + box.h + 3.4],
      [box.x + box.w * 0.2, box.y + box.h],
      [box.x, box.y + box.h],
    ],
    true,
  );

export const REVIEW_ICONS = {
  inspect: [
    ...doc({ x: 3, y: 2.4, w: 13, h: 18, fold: 3, lines: 2 }),
    ...magnifier({ cx: 15, cy: 14.8, r: 4.8, handle: 4.3, handleColor: PALETTE.info }),
  ],

  spelling: [
    letter('A', { height: 13.4, cx: 9.2, cy: 9.6 }),
    stroked(
      poly([
        [11.6, 16.8],
        [15, 20.4],
        [21.4, 12.4],
      ]),
      PALETTE.accent,
      'heavy',
    ),
  ],

  // A figure with its arms out: the standard accessibility mark.
  accessibility: [
    filled(circle(12, 4.6, 2.4), PALETTE.info),
    stroked(
      join('M3.6 9.4h16.8', 'M12 8.4v6.4', 'M12 14.8 8 21.4', 'M12 14.8 16 21.4'),
      PALETTE.ink,
      'heavy',
    ),
  ],

  // Two scripts, one becoming the other.
  translate: [
    outlined(rect(2.4, 3.4, 12, 12), PALETTE.paper, PALETTE.info, 'regular'),
    // A generic non-Latin mark: horizontal head stroke over two legs.
    stroked(
      join('M4.8 6.4h7.2', 'M8.4 5v1.4', 'M8.4 6.4v3', 'M8.4 9.4 5.8 13', 'M8.4 9.4 11 13'),
      PALETTE.info,
      'regular',
    ),
    outlined(rect(9.6, 8.6, 12, 12), PALETTE.paper, PALETTE.accent, 'regular'),
    letter('A', { height: 7.4, cx: 15.6, cy: 14.6, color: PALETTE.accent }),
  ],

  protect: [
    stroked('M8 10.4V7.8a4 4 0 0 1 8 0v2.6', PALETTE.ink, 'bold'),
    outlined(roundRect(4.4, 10.4, 15.2, 11.2, 1.8), PALETTE.warn, PALETTE.ink, 'regular'),
    filled(join(circle(12, 15.2, 1.9), rect(11.2, 15.2, 1.6, 3.8)), PALETTE.ink),
  ],

  commentAdd: [
    outlined(
      bubble({ x: 2.4, y: 4.4, w: 15.2, h: 10.6 }),
      PALETTE.warnSoft,
      PALETTE.warnDeep,
      'regular',
    ),
    ...bars({
      x: 5,
      y: 8,
      width: 10,
      widths: [1, 0.6],
      gap: 3,
      color: PALETTE.warnDeep,
      weight: 'regular',
    }),
    ...badge({ glyph: 'plus', corner: 'br' }),
  ],

  link: [
    chainLink({ from: [5.8, 18.2], to: [10.4, 13.6] }),
    chainLink({ from: [13.6, 10.4], to: [18.2, 5.8] }),
    stroked(line([10.2, 13.8], [13.8, 10.2]), PALETTE.info, 'bold'),
  ],
} satisfies Record<string, IconDefinition>;
