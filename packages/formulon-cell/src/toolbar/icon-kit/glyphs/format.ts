/** Font, border, fill-colour and number-format icons. */

import { letter, strokeLetter, strokeLetterWidth } from '../letters.js';
import { circle, ellipse, hRules, join, line, poly, rect, vRules } from '../path.js';
import { arrow, caret, filled, outlined, place, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** Colour swatch under a letterform, shared by the font-colour family. */
const swatch = (color: string) => filled(rect(3, 18.6, 18, 3), color);

/** Cap height used wherever a letterform sits above a swatch. */
const CAPPED = 13.6;

/** Baseline of the font-size letter, shared so grow and shrink sit on one line. */
const SIZE_BASELINE = 20.6;

/** The font-size pair: a light A with a caret at its shoulder. */
const sizeMark = (height: number, pointing: 'up' | 'down') => [
  strokeLetter('A', { height, cx: 9.4, cy: SIZE_BASELINE - height / 2, weight: 'bold' }),
  caret({ cx: 18.6, cy: 4.6, width: 4.6, pointing, color: PALETTE.info, weight: 'bold' }),
];

/** Digit height of the decimal-places marks. */
const DIGIT = 6.6;

/**
 * One row of a decimal-places mark: a point and `zeros` zeros, right-aligned
 * on `right`, with an optional arrow coming in from the left.
 */
const decimalRow = (options: {
  right: number;
  cy: number;
  zeros: number;
  arrow?: 'in' | 'out';
}) => {
  const digit = strokeLetterWidth('zero', DIGIT);
  const pitch = digit + 1.6;
  const firstCx = options.right - digit / 2 - pitch * (options.zeros - 1);
  const dotX = firstCx - digit / 2 - 1.4;
  const segments: IconDefinition[number][] = [
    filled(circle(dotX, options.cy + DIGIT / 2 - 0.6, 0.8), PALETTE.ink),
    ...Array.from({ length: options.zeros }, (_, i) =>
      strokeLetter('zero', { height: DIGIT, cx: firstCx + pitch * i, cy: options.cy }),
    ),
  ];
  if (options.arrow) {
    const tail = dotX - 7.6;
    const head = dotX - 1.4;
    segments.push(
      ...arrow({
        from: options.arrow === 'out' ? [head, options.cy] : [tail, options.cy],
        to: options.arrow === 'out' ? [tail, options.cy] : [head, options.cy],
        color: PALETTE.info,
        weight: 'regular',
      }),
    );
  }
  return segments;
};

/** Paint bucket, upright in its own box; tilted into place by the fill icon. */
const bucket = [
  stroked('M5.9 7.4a3.5 3.5 0 0 1 7 0', PALETTE.ink, 'thin'),
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
];

export const FORMAT_ICONS = {
  bold: [letter('B', { height: 17 })],
  italic: [strokeLetter('I', { height: 16.4, weight: 'regular' })],
  underline: [
    strokeLetter('U', { height: CAPPED, cy: 10, weight: 'regular' }),
    filled(rect(4.4, 19.6, 15.2, 1), PALETTE.ink),
  ],
  underlineSingle: [
    strokeLetter('U', { height: CAPPED, cy: 10, weight: 'regular' }),
    filled(rect(4.4, 19.6, 15.2, 1), PALETTE.ink),
  ],
  underlineDouble: [
    strokeLetter('U', { height: 12.4, cy: 9.2, weight: 'regular' }),
    filled(rect(4.4, 17.5, 15.2, 1), PALETTE.ink),
    filled(rect(4.4, 20.5, 15.2, 1), PALETTE.ink),
  ],
  strike: [
    letter('S', { height: 15.4, cy: 11.6 }),
    filled(rect(2.6, 11.2, 18.8, 1), PALETTE.danger),
  ],

  fontColor: [
    strokeLetter('A', { height: CAPPED, cy: 9.6, weight: 'bold' }),
    swatch(PALETTE.danger),
  ],
  // The bucket tips toward the swatch and the drop leaving its lip is the
  // pour; the colour it lays down is the swatch, not the bucket.
  fillColor: [
    ...place(bucket, { box: [3.2, 3.9, 12.4, 13.1], size: 13.2, cx: 10.4, cy: 9.8, rotate: 32 }),
    filled(
      join(
        circle(19.2, 12.8, 1.6),
        poly(
          [
            [19.2, 8.6],
            [17.6, 12.8],
            [20.8, 12.8],
          ],
          true,
        ),
      ),
      PALETTE.info,
    ),
    swatch(PALETTE.warn),
  ],
  fontGrow: sizeMark(15.2, 'up'),
  fontShrink: sizeMark(12, 'down'),

  // Dotted cell edges with one solid rule: the border about to be applied.
  borders: [
    stroked(
      join(rect(2.6, 2.6, 18.8, 18.8), hRules(2.6, 12, 18.8, 0, 1), vRules(12, 2.6, 18.8, 0, 1)),
      PALETTE.grid,
      'regular',
      { cap: 'butt', dash: '1 1.35' },
    ),
    stroked(line([2.6, 21.4], [21.4, 21.4]), PALETTE.ink, 'heavy', { cap: 'butt' }),
  ],

  // Accounting format: a banknote with the coin stack it is counted in.
  currency: [
    outlined(rect(2.4, 3.6, 15.2, 10.4), PALETTE.paper, PALETTE.ink, 'regular'),
    stroked(join('M5.2 6.4v4.8', 'M14.8 6.4v4.8'), PALETTE.ink, 'thin', { cap: 'butt' }),
    stroked(ellipse(10, 8.8, 2.6, 2.6), PALETTE.ink, 'thin'),
    ...[18.6, 16, 13.4].map((cy) =>
      outlined(
        join(ellipse(17.2, cy, 4.2, 1.6), `M13 ${cy}v1.2a4.2 1.6 0 0 0 8.4 0V${cy}`),
        PALETTE.paper,
        PALETTE.alt,
        'thin',
      ),
    ),
  ],
  percent: [
    stroked(circle(7, 7, 3.2), PALETTE.ink, 'regular'),
    stroked(circle(17, 17, 3.2), PALETTE.ink, 'regular'),
    stroked(line([18.6, 3.4], [5.4, 20.6]), PALETTE.ink, 'regular'),
  ],
  // The separator itself: a bowl with a tail that tapers, rather than a
  // wedge hanging off a disc.
  comma: [
    // Kept as two paths: merged into one, the tail's winding cancels the bowl
    // where they overlap and a hairline notch shows through.
    filled(circle(12, 7.7, 4.2), PALETTE.ink),
    filled('M15.9 9.1c.7 4.6-1.6 8.4-6.9 11.4 3.4-3.4 5-6.9 4.7-10.4z', PALETTE.ink),
  ],
  // Increase decimals: a place arrives in front of ".0" and it becomes ".00".
  decUp: [
    ...decimalRow({ right: 18.6, cy: 7.4, zeros: 1, arrow: 'out' }),
    ...decimalRow({ right: 18.6, cy: 16.6, zeros: 2 }),
  ],
  // Decrease decimals: ".00" loses a place on its way to ".0".
  decDown: [
    ...decimalRow({ right: 18.6, cy: 7.4, zeros: 2 }),
    ...decimalRow({ right: 18.6, cy: 16.6, zeros: 1, arrow: 'in' }),
  ],
} satisfies Record<string, IconDefinition>;
