/** Sort and filter icons. */

import { letter } from '../letters.js';
import { hLine, join } from '../path.js';
import { arrow, badge, funnel, headed, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** The sort key column: two glyphs stacked in the order they end up in. */
const sortKeys = (top: 'A' | 'Z', bottom: 'A' | 'Z') => [
  letter(top === 'A' ? 'aSmall' : 'Z', { height: 7.4, cx: 7, cy: 6.4, color: PALETTE.ink }),
  letter(bottom === 'A' ? 'aSmall' : 'Z', { height: 7.4, cx: 7, cy: 17.6, color: PALETTE.info }),
];

/** Direction arrow shared by the sort icons, always in the same column. */
const sortArrow = [
  ...arrow({ from: [17.6, 3.6], to: [17.6, 20.4], color: PALETTE.accent, head: 'solid' }),
];

export const SORT_FILTER_ICONS = {
  sortAsc: [...sortKeys('A', 'Z'), ...sortArrow],
  sortDesc: [...sortKeys('Z', 'A'), ...sortArrow],

  // Sort *and* filter: the cone with the direction arrow beside it.
  sortFilter: [
    ...funnel({ scale: 0.82 }),
    ...arrow({
      from: [19.4, 4.4],
      to: [19.4, 19.6],
      color: PALETTE.accent,
      head: 'solid',
      tail: 'solid',
      weight: 'regular',
    }),
  ],
  // Custom sort: the grid with a two-way arrow, no fixed direction.
  sortCustom: [
    ...sheet({ x: 2.4, y: 4.4, w: 13.2, h: 15.2, cols: 2, rows: 3 }),
    ...arrow({
      from: [19, 4.6],
      to: [19, 19.4],
      color: PALETTE.info,
      head: 'solid',
      tail: 'solid',
      weight: 'regular',
    }),
  ],

  filter: [...funnel()],
  filterToggle: [...funnel()],
  filterByValue: [
    ...funnel({ scale: 0.86 }),
    stroked(join(hLine(14.8, 17.8, 6.4), hLine(14.8, 20.8, 6.4)), PALETTE.accent, 'bold'),
  ],
  filterClear: [
    ...funnel({ scale: 0.8, frame: PALETTE.grid }),
    ...badge({ glyph: 'cross', corner: 'br' }),
  ],
  filterAdvanced: [
    ...funnel({ scale: 0.86 }),
    stroked(
      join(hLine(14.8, 15.4, 6.4), hLine(14.8, 18.2, 6.4), hLine(14.8, 21, 6.4)),
      PALETTE.grid,
      'regular',
    ),
  ],
  // Reapply: the cone with the refresh loop closing back on itself.
  filterReapply: [
    ...funnel({ scale: 0.8, frame: PALETTE.grid }),
    ...headed({
      shaft: 'M20.6 18.4a3.6 3.6 0 1 1-1.4-2.9',
      tip: [15.4, 15.4],
      headDir: [-0.9, 0.44],
      color: PALETTE.accent,
      weight: 'regular',
    }),
  ],
} satisfies Record<string, IconDefinition>;
