/** Fill-direction, series and autosum icons. */

import { letter } from '../letters.js';
import { poly } from '../path.js';
import { arrow, bars, filled, outlined, sheet } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/**
 * The fill family reads as "this grid, filled that way": a grid occupying the
 * side the fill comes from, and one arrow showing the direction. All five share
 * the same grid box so they line up with each other in a menu.
 */
const GRID = { x: 2.4, y: 2.4, w: 13.2, h: 19.2 } as const;
const GRID_WIDE = { x: 2.4, y: 2.4, w: 19.2, h: 13.2 } as const;

const directionalFill = (
  box: typeof GRID | typeof GRID_WIDE,
  cols: number,
  rows: number,
  from: readonly [number, number],
  to: readonly [number, number],
) => [...sheet({ ...box, cols, rows }), ...arrow({ from, to, color: PALETTE.info, head: 'solid' })];

export const FILL_ICONS = {
  autosum: [letter('sigma', { height: 18.4, color: PALETTE.info })],

  // The Fill menu's own icon: the source row spreading through the block.
  fill: [
    ...sheet({
      x: 2.4,
      y: 2.4,
      w: 19.2,
      h: 19.2,
      cols: 3,
      rows: 3,
      bands: [{ axis: 'row', index: 0, fill: PALETTE.accentSoft }],
    }),
    ...arrow({
      from: [12, 8.4],
      to: [12, 19.2],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
  ],
  fillDown: directionalFill(GRID, 2, 3, [19, 5.4], [19, 18.6]),
  fillUp: directionalFill(GRID, 2, 3, [19, 18.6], [19, 5.4]),
  fillRight: directionalFill(GRID_WIDE, 3, 2, [5.4, 19], [18.6, 19]),
  fillLeft: directionalFill(GRID_WIDE, 3, 2, [18.6, 19], [5.4, 19]),

  // Fill across grouped sheets: a stack, with the fill landing on all of them.
  fillGroup: [
    ...sheet({
      x: 6.4,
      y: 2.4,
      w: 15.2,
      h: 15.2,
      cols: 2,
      rows: 2,
      surface: PALETTE.mute,
      frame: PALETTE.grid,
    }),
    ...sheet({ x: 2.4, y: 6.4, w: 15.2, h: 15.2, cols: 2, rows: 2 }),
  ],

  // Fill series: the values step up as they go down the column.
  fillSeries: [
    ...sheet({ x: 2.4, y: 2.4, w: 12, h: 19.2, cols: 1, rows: 3 }),
    ...bars({
      x: 4.8,
      y: 6.4,
      width: 7.2,
      widths: [0.45, 0.7, 1],
      gap: 6.4,
      color: PALETTE.info,
      weight: 'regular',
    }),
    ...arrow({ from: [18.4, 5.4], to: [18.4, 18.6], color: PALETTE.accent, head: 'solid' }),
  ],

  // Justify redistributes one block of text across the rows below it.
  fillJustify: [
    ...sheet({ x: 2.4, y: 2.4, w: 19.2, h: 19.2, cols: 1, rows: 3, rules: PALETTE.gridLight }),
    ...bars({
      x: 5,
      y: 6,
      width: 14,
      widths: [1, 1, 0.6],
      gap: 6.4,
      color: PALETTE.ink,
      weight: 'regular',
    }),
  ],

  // Flash fill: the grid, with the pattern recognised in a stroke.
  flashFill: [
    ...sheet({ x: 2.4, y: 3.4, w: 19.2, h: 17.2, cols: 3, rows: 2 }),
    filled(
      poly(
        [
          [14.6, 2.6],
          [9, 12.4],
          [12.4, 12.4],
          [10.6, 21.4],
          [16.6, 11],
          [13, 11],
        ],
        true,
      ),
      PALETTE.paper,
    ),
    outlined(
      poly(
        [
          [14.6, 2.6],
          [9, 12.4],
          [12.4, 12.4],
          [10.6, 21.4],
          [16.6, 11],
          [13, 11],
        ],
        true,
      ),
      PALETTE.alt,
      PALETTE.altDeep,
      'hairline',
    ),
  ],
} satisfies Record<string, IconDefinition>;
