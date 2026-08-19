/** Find, replace and go-to icons. */

import { letter } from '../letters.js';
import { join, line, poly } from '../path.js';
import {
  arrow,
  badge,
  bars,
  headed,
  magnifier,
  outlined,
  place,
  sheet,
  stroked,
} from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/**
 * The find family is a lens over a subject. The lens keeps one position and
 * size across every variant so the set reads as one family, and the subject
 * changes underneath it.
 */
const LENS = { cx: 15.4, cy: 8.6, r: 5.4, handle: 5.6 } as const;

const overLens = (subject: readonly IconDefinition[number][]) => [
  ...subject,
  ...magnifier({ ...LENS, surface: PALETTE.paper, handleColor: PALETTE.info }),
];

/** The card the find-* variants inspect. */
const card = (fill: string) => [
  outlined(
    poly(
      [
        [2.4, 5],
        [14.6, 5],
        [14.6, 19.6],
        [2.4, 19.6],
      ],
      true,
    ),
    fill,
    PALETTE.grid,
    'regular',
  ),
];

/** Lowercase f with a cross — the formula mark, reused from the paste family. */
const formulaMark = place(
  [
    stroked(join('M12.9 6.4c-1.9-.6-3 .3-3.4 2.3l-2 9.6', 'M7.9 10.4h4.4'), PALETTE.violet, 'bold'),
    stroked(
      join(line([13.1, 13.4], [16.3, 17.2]), line([16.3, 13.4], [13.1, 17.2])),
      PALETTE.ink,
      'regular',
    ),
  ],
  { box: [6.9, 6.4, 9.4, 10.8], size: 8.6, cx: 7.6, cy: 13.4 },
);

export const FIND_ICONS = {
  find: [...magnifier({ cx: 9.8, cy: 9.8, r: 7, handle: 6.4, handleColor: PALETTE.info })],

  // Replace: the lens over text that is being swapped for other text.
  replaceFind: [
    ...bars({ x: 2.4, y: 5.4, width: 10.4, widths: [1, 0.7], gap: 3.4, color: PALETTE.ink }),
    ...headed({
      shaft: 'M3.4 14.4h7.4a2.4 2.4 0 0 1 0 4.8H4.6',
      tip: [3.4, 19.2],
      headDir: [-1, 0],
      color: PALETTE.accent,
      weight: 'regular',
    }),
    ...magnifier({ ...LENS, surface: PALETTE.paper, handleColor: PALETTE.info }),
  ],

  findFormulas: [...overLens([...card(PALETTE.paper), ...formulaMark])],
  findConstants: [
    ...overLens([
      ...card(PALETTE.paper),
      letter('digits', { height: 7.4, cx: 7.6, cy: 14, color: PALETTE.accent }),
    ]),
  ],
  findConditional: [
    ...overLens([
      ...sheet({
        x: 2.4,
        y: 5,
        w: 12.2,
        h: 14.6,
        cols: 1,
        rows: 3,
        frame: PALETTE.grid,
        bands: [
          { axis: 'row', index: 0, fill: PALETTE.scaleHigh },
          { axis: 'row', index: 1, fill: PALETTE.scaleMid },
          { axis: 'row', index: 2, fill: PALETTE.scaleLow },
        ],
        rules: PALETTE.paper,
      }),
    ]),
  ],
  findValidation: [
    ...overLens([
      ...card(PALETTE.paper),
      ...bars({
        x: 4.8,
        y: 9.4,
        width: 7.4,
        widths: [1, 0.7],
        gap: 3.2,
        color: PALETTE.grid,
        weight: 'regular',
      }),
      stroked(
        poly([
          [5, 15.6],
          [7.2, 17.8],
          [12.2, 12.4],
        ]),
        PALETTE.accent,
        'bold',
      ),
    ]),
  ],
  findComments: [
    ...overLens([
      outlined(
        poly(
          [
            [2.4, 5],
            [14.6, 5],
            [14.6, 15],
            [8.2, 15],
            [5, 19.4],
            [5, 15],
            [2.4, 15],
          ],
          true,
        ),
        PALETTE.warnSoft,
        PALETTE.warnDeep,
        'regular',
      ),
    ]),
  ],

  // Go to: a jump into a specific cell.
  goTo: [
    ...sheet({ x: 2.4, y: 3.4, w: 12.6, h: 17.2, cols: 2, rows: 3 }),
    ...arrow({ from: [11.6, 12], to: [21.2, 12], color: PALETTE.accent, head: 'solid' }),
  ],
  goToSpecial: [
    ...sheet({ x: 2.4, y: 3.4, w: 15.2, h: 17.2, cols: 2, rows: 3 }),
    ...badge({ glyph: 'star', corner: 'br', tone: PALETTE.violet }),
  ],
} satisfies Record<string, IconDefinition>;
