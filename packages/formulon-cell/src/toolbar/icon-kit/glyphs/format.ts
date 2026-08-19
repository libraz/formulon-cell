/** Font, border, fill-colour and number-format icons. */

import { letter } from '../letters.js';
import { circle, join, line, poly, rect } from '../path.js';
import { arrow, filled, outlined, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** Colour swatch under a letterform, shared by the font-colour family. */
const swatch = (color: string) => filled(rect(3, 18.6, 18, 3), color);

/** Cap height used wherever a letterform sits above a swatch. */
const CAPPED = 13.6;

export const FORMAT_ICONS = {
  bold: [letter('B', { height: 17.6 })],
  italic: [letter('I', { height: 17.6 })],
  underline: [letter('U', { height: CAPPED, cy: 9.8 }), swatch(PALETTE.ink)],
  underlineSingle: [letter('U', { height: CAPPED, cy: 9.8 }), swatch(PALETTE.ink)],
  underlineDouble: [
    letter('U', { height: 12.4, cy: 9 }),
    filled(rect(3, 17, 18, 1.9), PALETTE.ink),
    filled(rect(3, 20, 18, 1.9), PALETTE.ink),
  ],
  strike: [
    letter('S', { height: 15.4, cy: 11.6 }),
    filled(rect(2.6, 10.4, 18.8, 2.6), PALETTE.danger),
  ],

  fontColor: [letter('A', { height: CAPPED, cy: 9.8 }), swatch(PALETTE.danger)],
  fillColor: [
    // A bucket over the swatch, mid-pour. The handle arcs clear of the rim so
    // the silhouette cannot be read as a bag.
    stroked('M5.9 7.4a3.5 3.5 0 0 1 7 0', PALETTE.grid, 'thin'),
    outlined(
      poly(
        [
          [3.2, 7.4],
          [15.6, 7.4],
          [13.4, 17],
          [5.4, 17],
        ],
        true,
      ),
      PALETTE.paper,
      PALETTE.ink,
      'regular',
    ),
    filled(
      poly(
        [
          [4.3, 12.2],
          [14.5, 12.2],
          [13.4, 17],
          [5.4, 17],
        ],
        true,
      ),
      PALETTE.warn,
    ),
    // The drop leaving the lip ties the bucket to the colour it lays down.
    filled(
      join(
        circle(18.8, 12.6, 1.8),
        poly(
          [
            [18.8, 8],
            [17, 12.6],
            [20.6, 12.6],
          ],
          true,
        ),
      ),
      PALETTE.warn,
    ),
    swatch(PALETTE.warn),
  ],
  fontGrow: [
    letter('A', { height: 14.6, cx: 8, cy: 12 }),
    ...arrow({ from: [18.2, 19], to: [18.2, 4.6], color: PALETTE.accent, head: 'solid' }),
  ],
  fontShrink: [
    letter('A', { height: 14.6, cx: 8, cy: 12 }),
    ...arrow({ from: [18.2, 4.6], to: [18.2, 19], color: PALETTE.accent, head: 'solid' }),
  ],

  borders: [...sheet({ x: 2.4, y: 2.4, w: 19.2, h: 19.2, cols: 3, rows: 3, rules: PALETTE.grid })],

  currency: [
    outlined(circle(12, 12, 9.2), PALETTE.warnSoft, PALETTE.warnDeep, 'regular'),
    stroked(
      join('M8.4 7.6 12 12l3.6-4.4', 'M12 12v4.8', 'M8.8 12.6h6.4', 'M8.8 14.8h6.4'),
      PALETTE.accentDeep,
      'bold',
    ),
  ],
  percent: [
    outlined(circle(7.2, 7.2, 3.7), PALETTE.paper, PALETTE.ink, 'bold'),
    outlined(circle(16.8, 16.8, 3.7), PALETTE.paper, PALETTE.ink, 'bold'),
    stroked(line([19.2, 3.8], [4.8, 20.2]), PALETTE.ink, 'heavy'),
  ],
  // The separator itself: a bowl with a tail that tapers, rather than a
  // wedge hanging off a disc.
  comma: [
    // Kept as two paths: merged into one, the tail's winding cancels the bowl
    // where they overlap and a hairline notch shows through.
    filled(circle(12, 8.8, 4.2), PALETTE.ink),
    filled('M15.9 10.2c.7 4.6-1.6 8.4-6.9 11.4 3.4-3.4 5-6.9 4.7-10.4z', PALETTE.ink),
  ],
  decUp: [
    ...decimalMark(),
    ...arrow({ from: [18.6, 18.4], to: [18.6, 5.6], color: PALETTE.accent, head: 'solid' }),
  ],
  decDown: [
    ...decimalMark(),
    ...arrow({ from: [18.6, 5.6], to: [18.6, 18.4], color: PALETTE.accent, head: 'solid' }),
  ],
} satisfies Record<string, IconDefinition>;

/** The ".00" left half shared by the two decimal-place icons. */
function decimalMark() {
  return [
    filled(circle(3.2, 16.6, 1.5), PALETTE.ink),
    outlined(circle(7.8, 12.6, 2.9), 'none', PALETTE.ink, 'bold'),
    outlined(circle(14.2, 12.6, 2.9), 'none', PALETTE.ink, 'bold'),
  ];
}
