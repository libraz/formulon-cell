/** Table, pivot-table and cell-style icons. */

import { rect, roundRect } from '../path.js';
import { badge, bars, filled, headed, outlined, sheet } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** A table: a grid with its header row filled in. */
const table = (box: { x: number; y: number; w: number; h: number }, header: string) => [
  ...sheet({
    ...box,
    cols: 3,
    rows: 4,
    bands: [{ axis: 'row', index: 0, fill: header }],
  }),
];

const FULL = { x: 2.4, y: 3.4, w: 19.2, h: 17.2 } as const;
const BADGED = { x: 2.4, y: 4.4, w: 16.2, h: 16.2 } as const;

export const TABLE_ICONS = {
  table: table(FULL, PALETTE.accent),
  tableStyle: [
    ...table(BADGED, PALETTE.info),
    ...badge({ glyph: 'pencil', corner: 'br', tone: PALETTE.alt }),
  ],

  // A pivot pulls a summary out of the source grid.
  pivotTable: [
    ...sheet({ x: 2.4, y: 2.4, w: 11.4, h: 11.4, cols: 2, rows: 2, frame: PALETTE.grid }),
    ...sheet({
      x: 10.6,
      y: 10.6,
      w: 11,
      h: 11,
      cols: 2,
      rows: 2,
      bands: [{ axis: 'row', index: 0, fill: PALETTE.accentSoft }],
    }),
  ],
  pivotRecommended: [
    ...sheet({
      x: 2.4,
      y: 4.4,
      w: 16.2,
      h: 16.2,
      cols: 3,
      rows: 3,
      bands: [{ axis: 'row', index: 0, fill: PALETTE.infoSoft }],
    }),
    ...badge({ glyph: 'star', corner: 'tr', tone: PALETTE.warn, onTone: PALETTE.warnDeep }),
  ],
  // Placed on a sheet that already exists: the summary lands in the target.
  pivotExistingSheet: [
    ...sheet({ x: 2.4, y: 2.4, w: 10.4, h: 10.4, cols: 2, rows: 2, frame: PALETTE.grid }),
    ...sheet({
      x: 11.6,
      y: 11.6,
      w: 10,
      h: 10,
      cols: 2,
      rows: 2,
      bands: [{ axis: 'row', index: 0, fill: PALETTE.infoSoft }],
    }),
    ...headed({
      shaft: 'M8.6 13.4v3.4',
      tip: [8.6, 19.4],
      headDir: [0, 1],
      color: PALETTE.info,
      weight: 'regular',
    }),
  ],

  // Format cells: the grid with the dialog's own swatch beside it.
  formatCells: [
    ...sheet({ x: 2.4, y: 4.4, w: 16.2, h: 16.2, cols: 3, rows: 3 }),
    ...badge({ glyph: 'pencil', corner: 'br', tone: PALETTE.info }),
  ],
  conditional: [
    ...sheet({ x: 2.4, y: 3.4, w: 12.2, h: 17.2, cols: 1, rows: 3 }),
    filled(rect(16.4, 4.4, 5.2, 4.4), PALETTE.scaleLow),
    filled(rect(16.4, 9.8, 5.2, 4.4), PALETTE.scaleMid),
    filled(rect(16.4, 15.2, 5.2, 4.4), PALETTE.scaleHigh),
  ],
  cellStyles: [
    outlined(roundRect(2.4, 4.4, 8.6, 6.4, 1.2), PALETTE.accentSoft, PALETTE.accent, 'regular'),
    outlined(roundRect(13, 4.4, 8.6, 6.4, 1.2), PALETTE.warnSoft, PALETTE.warnDeep, 'regular'),
    outlined(roundRect(2.4, 13.2, 8.6, 6.4, 1.2), PALETTE.infoSoft, PALETTE.info, 'regular'),
    outlined(roundRect(13, 13.2, 8.6, 6.4, 1.2), PALETTE.dangerSoft, PALETTE.danger, 'regular'),
    ...bars({
      x: 4.4,
      y: 7.6,
      width: 4.6,
      widths: [1],
      gap: 0,
      color: PALETTE.accent,
      weight: 'regular',
    }),
    ...bars({
      x: 15,
      y: 7.6,
      width: 4.6,
      widths: [1],
      gap: 0,
      color: PALETTE.warnDeep,
      weight: 'regular',
    }),
    ...bars({
      x: 4.4,
      y: 16.4,
      width: 4.6,
      widths: [1],
      gap: 0,
      color: PALETTE.info,
      weight: 'regular',
    }),
    ...bars({
      x: 15,
      y: 16.4,
      width: 4.6,
      widths: [1],
      gap: 0,
      color: PALETTE.danger,
      weight: 'regular',
    }),
  ],
} satisfies Record<string, IconDefinition>;
