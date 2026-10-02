/** Formula auditing, named ranges, validation and calculation icons. */

import { letter } from '../letters.js';
import { circle, ellipse, join, line, poly, rect, roundRect } from '../path.js';
import {
  arrow,
  badge,
  bars,
  filled,
  headed,
  magnifier,
  outlined,
  place,
  ring,
  sheet,
  star,
  stroked,
} from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition, IconSegment } from '../types.js';

/** Shared Office-style book frame used by the Function Library families. */
const functionFamilyBook = (tone: string, mark: readonly IconSegment[]): IconDefinition => [
  outlined(roundRect(3, 2.8, 18, 18.4, 1.5), PALETTE.paper, tone, 'regular'),
  stroked(line([6.8, 4.2], [6.8, 19.8]), tone, 'thin'),
  ...mark,
];

const functionFamilyIcons = {
  functionRecent: functionFamilyBook(PALETTE.info, [
    filled(star(14.3, 11.8, 4.4, 2), PALETTE.info),
  ]),
  functionFinancial: functionFamilyBook(PALETTE.accent, [
    filled(ellipse(13.5, 8.7, 3.2, 1.5), PALETTE.accent),
    filled(ellipse(14.5, 12, 3.2, 1.5), PALETTE.accent),
    filled(ellipse(15.5, 15.3, 3.2, 1.5), PALETTE.accent),
    stroked(
      join(line([10.3, 8.7], [10.3, 15.3]), line([18.7, 8.7], [18.7, 15.3])),
      PALETTE.accent,
      'thin',
    ),
  ]),
  functionLogical: functionFamilyBook(PALETTE.violet, [
    stroked(
      'M11.1 10.2c0-2 1.5-3.4 3.5-3.4 2 0 3.4 1.2 3.4 3 0 2.2-2.8 2.6-3.2 4.4',
      PALETTE.violet,
      'regular',
    ),
    filled(ellipse(14.7, 17.1, 1, 1), PALETTE.violet),
  ]),
  functionText: functionFamilyBook(PALETTE.info, [
    letter('A', { height: 9, cx: 14.4, cy: 12, color: PALETTE.info }),
  ]),
  functionDateTime: functionFamilyBook(PALETTE.danger, [
    outlined(circle(14.4, 11.5, 4.5), PALETTE.paper, PALETTE.danger, 'regular'),
    stroked(
      join(line([14.4, 11.5], [14.4, 8.6]), line([14.4, 11.5], [17, 13])),
      PALETTE.danger,
      'thin',
    ),
  ]),
  functionLookup: functionFamilyBook(PALETTE.info, [
    ...magnifier({
      cx: 13.5,
      cy: 10.8,
      r: 4.1,
      handle: 3.8,
      frame: PALETTE.info,
      handleColor: PALETTE.info,
    }),
  ]),
  functionMath: functionFamilyBook(PALETTE.accent, [
    stroked(ellipse(14.2, 11.8, 4.1, 4.8), PALETTE.accent, 'regular'),
    stroked(line([10.1, 11.8], [18.3, 11.8]), PALETTE.accent, 'thin'),
  ]),
  functionMore: functionFamilyBook(PALETTE.danger, [
    filled(ellipse(11.5, 12, 1, 1), PALETTE.danger),
    filled(ellipse(14.5, 12, 1, 1), PALETTE.danger),
    filled(ellipse(17.5, 12, 1, 1), PALETTE.danger),
  ]),
} satisfies Record<string, IconDefinition>;

/** The card the validation icons annotate. */
const card = (fill = PALETTE.paper) => [
  outlined(rect(3.4, 3.4, 15.2, 17.2), fill, PALETTE.ink, 'regular'),
];

/** Cell boxes wired together, for the precedent/dependent tracers. */
const traceCells = () => [
  outlined(rect(2.4, 8.6, 6.8, 6.8), PALETTE.paper, PALETTE.ink, 'regular'),
  outlined(rect(14.8, 2.6, 6.8, 6.8), PALETTE.paper, PALETTE.grid, 'regular'),
  outlined(rect(14.8, 14.6, 6.8, 6.8), PALETTE.paper, PALETTE.grid, 'regular'),
];

export const FORMULA_ICONS = {
  ...functionFamilyIcons,
  // Precedents flow into the active cell; dependents flow out of it.
  trace: [
    ...traceCells(),
    ...headed({
      shaft: 'M14.8 6h-3.2v6',
      tip: [11.6, 12],
      headDir: [0, 1],
      color: PALETTE.info,
      weight: 'regular',
    }),
    ...headed({
      shaft: 'M14.8 18h-3.2v-6',
      tip: [11.6, 12],
      headDir: [0, -1],
      color: PALETTE.info,
      weight: 'regular',
    }),
  ],
  dependents: [
    ...traceCells(),
    ...headed({
      shaft: 'M11.6 12V6h3.2',
      tip: [14.8, 6],
      headDir: [1, 0],
      color: PALETTE.accent,
      weight: 'regular',
    }),
    ...headed({
      shaft: 'M11.6 12v6h3.2',
      tip: [14.8, 18],
      headDir: [1, 0],
      color: PALETTE.accent,
      weight: 'regular',
    }),
  ],
  clearArrows: [
    ...traceCells(),
    ...headed({
      shaft: 'M11.6 12V6h3.2',
      tip: [14.8, 6],
      headDir: [1, 0],
      color: PALETTE.grid,
      weight: 'regular',
    }),
    ...headed({
      shaft: 'M11.6 12v6h3.2',
      tip: [14.8, 18],
      headDir: [1, 0],
      color: PALETTE.grid,
      weight: 'regular',
    }),
    ...badge({ glyph: 'cross', corner: 'br' }),
  ],

  errorChecking: [
    ...card(),
    ...bars({
      x: 6,
      y: 7.6,
      width: 10,
      widths: [1, 0.7],
      gap: 3.4,
      color: PALETTE.grid,
      weight: 'regular',
    }),
    ...badge({ glyph: 'bang', corner: 'br', onTone: PALETTE.ink }),
  ],
  calcOptions: [
    ...sheet({ x: 2.4, y: 4.4, w: 16.2, h: 16.2, cols: 2, rows: 3 }),
    ...badge({ glyph: 'dots', corner: 'br', tone: PALETTE.info }),
  ],
  // Watch window: an eye over the value being tracked.
  watch: [
    stroked(
      'M2 12c3.4-4.6 6.8-6.9 10-6.9S18.6 7.4 22 12c-3.4 4.6-6.8 6.9-10 6.9S5.4 16.6 2 12z',
      PALETTE.ink,
      'regular',
    ),
    outlined(circle(12, 12, 3.4), PALETTE.paper, PALETTE.accent, 'bold'),
    filled(circle(12, 12, 1.4), PALETTE.accent),
  ],

  // The fx mark on its own, at full optical size.
  function: [
    ...place(
      [
        stroked(
          join('M12.9 6.4c-1.9-.6-3 .3-3.4 2.3l-2 9.6', 'M7.9 10.4h4.4'),
          PALETTE.info,
          'bold',
        ),
        stroked(
          join(line([13.1, 13.4], [16.3, 17.2]), line([16.3, 13.4], [13.1, 17.2])),
          PALETTE.ink,
          'regular',
        ),
      ],
      // Include the italic descender and round stroke caps in the source box.
      { box: [6.6, 5.3, 10.5, 13.9], size: 20 },
    ),
  ],

  dataValidation: [
    ...card(),
    ...bars({
      x: 6,
      y: 7.6,
      width: 10,
      widths: [1, 0.7],
      gap: 3.4,
      color: PALETTE.grid,
      weight: 'regular',
    }),
    ...badge({ glyph: 'check', corner: 'br' }),
  ],
  dataValidationCircle: [...card(), ring(11, 12, 6.6, 4.2, PALETTE.danger, 'heavy')],
  dataValidationClearCircles: [
    ...card(),
    ring(10.4, 12, 6, 3.8, PALETTE.grid, 'bold'),
    ...badge({ glyph: 'cross', corner: 'br' }),
  ],
  dataValidationClearRules: [
    ...card(),
    ...bars({
      x: 6,
      y: 7.6,
      width: 10,
      widths: [1, 0.7],
      gap: 3.4,
      color: PALETTE.grid,
      weight: 'regular',
    }),
    ...badge({ glyph: 'cross', corner: 'br' }),
  ],

  // A named range: the grid with its name tag attached.
  names: [
    ...sheet({ x: 2.4, y: 6.6, w: 19.2, h: 15, cols: 3, rows: 2 }),
    filled(
      poly(
        [
          [2.4, 2.2],
          [13.4, 2.2],
          [16.4, 5.2],
          [13.4, 8.2],
          [2.4, 8.2],
        ],
        true,
      ),
      PALETTE.accent,
    ),
  ],
  namesCreateTop: nameFromEdge('row', 0),
  namesCreateBottom: nameFromEdge('row', 2),
  namesCreateLeft: nameFromEdge('col', 0),
  namesCreateRight: nameFromEdge('col', 2),
} satisfies Record<string, IconDefinition>;

/** Grid with the edge band that supplies the names highlighted. */
function nameFromEdge(axis: 'row' | 'col', index: number): IconDefinition {
  const band = { axis, index, fill: PALETTE.accent } as const;
  const arrowSpec =
    axis === 'row'
      ? ({
          from: [12, index === 0 ? 8 : 16] as const,
          to: [12, index === 0 ? 14.4 : 9.6] as const,
        } as const)
      : ({
          from: [index === 0 ? 8 : 16, 12] as const,
          to: [index === 0 ? 14.4 : 9.6, 12] as const,
        } as const);
  return [
    ...sheet({ x: 2.4, y: 2.4, w: 19.2, h: 19.2, cols: 3, rows: 3, bands: [band] }),
    ...arrow({
      from: arrowSpec.from,
      to: arrowSpec.to,
      color: PALETTE.info,
      head: 'solid',
      weight: 'regular',
    }),
  ];
}
