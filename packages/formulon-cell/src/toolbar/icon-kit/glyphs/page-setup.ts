/** Page Setup dialog art: the orientation choices and the margins preview page. */

import { hLine, join, vLine, vRules } from '../path.js';
import { arrow, doc, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** A page beside an upward arrow; landscape text runs down the page. */
const orientation = (landscape: boolean): IconDefinition => [
  ...arrow({ from: [4.5, 21], to: [4.5, 3], color: PALETTE.info, head: 'solid' }),
  ...doc({ x: 8, y: 2, w: 13, h: 20, fold: 3.5, lines: landscape ? 0 : 5 }),
  ...(landscape ? [stroked(vRules(11.5, 7, 12, 2.6, 3), PALETTE.grid, 'thin')] : []),
];

/** The margins preview: a sheet page with header and footer rules and margin ticks. */
const marginPreview: IconDefinition = [
  ...sheet({ x: 5, y: 2, w: 14, h: 20, cols: 4, rows: 9, frame: PALETTE.ink }),
  stroked(join(hLine(5, 4, 14), hLine(5, 20, 14)), PALETTE.ink, 'thin'),
  stroked(join(vLine(4, 4, 1.6), vLine(20, 4, 1.6)), PALETTE.ink, 'thin'),
];

export const PAGE_SETUP_ICONS = {
  pageOrientationPortrait: orientation(false),
  pageOrientationLandscape: orientation(true),
  pageMarginsPreview: marginPreview,
} satisfies Record<string, IconDefinition>;
