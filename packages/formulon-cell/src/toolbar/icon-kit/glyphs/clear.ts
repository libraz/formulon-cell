/** Clear-* icons: one eraser body, one badge saying what is being cleared. */

import { letter } from '../letters.js';
import { hLine, join, line, poly } from '../path.js';
import { badge, bars, chainLink, outlined, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/**
 * The eraser, tipped over a rule it has just wiped. Sized so the badge in the
 * top-right corner never touches it.
 */
const eraser = (body: string, streak: string) => [
  outlined(
    poly(
      [
        [7.6, 4.2],
        [18.4, 9.8],
        [13.8, 16.4],
        [3, 10.8],
      ],
      true,
    ),
    body,
    PALETTE.ink,
    'regular',
  ),
  stroked(line([13, 7], [8.4, 13.6]), streak, 'thin'),
  stroked(hLine(2.6, 20.4, 18.8), PALETTE.info, 'bold'),
];

export const CLEAR_ICONS = {
  clear: eraser(PALETTE.paper, PALETTE.grid),
  clearAll: [...eraser(PALETTE.violetSoft, PALETTE.violet), ...badge({ glyph: 'cross' })],
  clearFormats: [
    ...eraser(PALETTE.violetSoft, PALETTE.violet),
    letter('A', { height: 7.6, cx: 18.4, cy: 5.4, color: PALETTE.violet }),
  ],
  clearContents: [
    ...sheet({ x: 2.4, y: 3.6, w: 16.4, h: 16.4, cols: 1, rows: 1 }),
    ...bars({
      x: 5,
      y: 8,
      width: 11.2,
      widths: [1, 0.75, 0.9],
      gap: 3.6,
      color: PALETTE.ink,
      weight: 'regular',
    }),
    ...badge({ glyph: 'cross' }),
  ],
  clearComments: [
    outlined(
      poly(
        [
          [2.4, 4],
          [17.6, 4],
          [17.6, 14.4],
          [9.6, 14.4],
          [5.6, 19.2],
          [5.6, 14.4],
          [2.4, 14.4],
        ],
        true,
      ),
      PALETTE.warnSoft,
      PALETTE.warnDeep,
      'regular',
    ),
    ...bars({
      x: 5,
      y: 7.6,
      width: 10,
      widths: [1, 0.6],
      gap: 3,
      color: PALETTE.warnDeep,
      weight: 'regular',
    }),
    ...badge({ glyph: 'cross' }),
  ],
  clearHyperlinks: [
    chainLink({ from: [5.6, 17], to: [9.8, 12.8] }),
    chainLink({ from: [14.2, 8.4], to: [18.4, 4.2] }),
    stroked(
      join(line([9.4, 15.4], [11.4, 13.4]), line([12.6, 12.6], [14.6, 10.6])),
      PALETTE.ink,
      'regular',
    ),
    ...badge({ glyph: 'cross', corner: 'br' }),
  ],
  clearConditional: [
    ...sheet({
      x: 2.4,
      y: 3.6,
      w: 16.4,
      h: 16.4,
      cols: 1,
      rows: 3,
      bands: [
        { axis: 'row', index: 0, fill: PALETTE.scaleHigh },
        { axis: 'row', index: 1, fill: PALETTE.scaleMid },
        { axis: 'row', index: 2, fill: PALETTE.scaleLow },
      ],
      rules: PALETTE.paper,
    }),
    ...badge({ glyph: 'cross' }),
  ],
} satisfies Record<string, IconDefinition>;
