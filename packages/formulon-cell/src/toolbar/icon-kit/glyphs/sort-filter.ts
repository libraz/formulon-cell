/** Sort and filter icons. */

import { letter, strokeLetter } from '../letters.js';
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

/** The 16px sort key of the desktop's filter panels: a capital A in info over
 *  or under an ink Z, the A keeping its colour in either order. */
const smallSortKeys = (top: 'A' | 'Z') => [
  letter('A', { height: 9.7, cx: 5.6, cy: top === 'A' ? 5.6 : 19.1, color: PALETTE.info }),
  letter('Z', { height: 9.7, cx: 5.6, cy: top === 'A' ? 19.1 : 5.6, color: PALETTE.ink }),
];

/** Short ink arrow beside the small sort key, its head carrying the icon. */
const smallSortArrow = [
  ...arrow({
    from: [19.5, 6.8],
    to: [19.5, 20.2],
    color: PALETTE.ink,
    head: 'solid',
    weight: 'heavy',
    headScale: 1.95,
  }),
];

export const SORT_FILTER_ICONS = {
  sortAsc: [...sortKeys('A', 'Z'), ...sortArrow],
  sortDesc: [...sortKeys('Z', 'A'), ...sortArrow],
  sortAscSmall: [...smallSortKeys('A'), ...smallSortArrow],
  sortDescSmall: [...smallSortKeys('Z'), ...smallSortArrow],

  // Sort *and* filter: the A-to-Z key beside the cone.
  sortFilter: [
    strokeLetter('A', { height: 7.6, cx: 5.6, cy: 6.6, color: PALETTE.info }),
    strokeLetter('Z', { height: 7.6, cx: 5.6, cy: 17.4 }),
    ...funnel({ scale: 0.62, cx: 16 }),
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
