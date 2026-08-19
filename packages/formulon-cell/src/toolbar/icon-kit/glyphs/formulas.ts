/** Formula auditing, named ranges, validation and calculation icons. */

import { circle, join, line, poly, rect } from '../path.js';
import {
  arrow,
  badge,
  bars,
  filled,
  headed,
  outlined,
  place,
  ring,
  sheet,
  stroked,
} from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

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
      { box: [6.9, 6.4, 9.4, 10.8], size: 18.4 },
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
