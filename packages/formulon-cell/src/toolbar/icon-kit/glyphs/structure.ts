/** Row/column/cell insertion, deletion, outlining and splitting. */

import { arrow, badge, bars, type SheetBands, sheet, stroked } from '../primitives.js';
import { OPTICAL_BADGED, PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/**
 * Grid box used by every badged structure icon. The badge overlaps this box's
 * corner rather than pushing it aside, so the grid stays centred.
 */
const BADGED_GRID = {
  x: 12 - OPTICAL_BADGED / 2,
  y: 12 - OPTICAL_BADGED / 2,
  w: OPTICAL_BADGED,
  h: OPTICAL_BADGED,
  cols: 3,
  rows: 3,
} as const;

/** Grid with the affected row, column or cell tinted. */
const targeted = (axis: 'row' | 'col' | 'cell', fill: string): SheetBands =>
  axis === 'cell'
    ? [
        { axis: 'row', index: 1, fill },
        { axis: 'col', index: 1, fill },
      ]
    : [{ axis, index: 1, fill }];

const structureIcon = (axis: 'row' | 'col' | 'cell', kind: 'add' | 'remove') => [
  ...sheet({
    ...BADGED_GRID,
    bands: targeted(axis, kind === 'add' ? PALETTE.accentSoft : PALETTE.dangerSoft),
  }),
  ...badge({ glyph: kind === 'add' ? 'plus' : 'cross' }),
];

export const STRUCTURE_ICONS = {
  insertRows: structureIcon('row', 'add'),
  insertCols: structureIcon('col', 'add'),
  insertCells: structureIcon('cell', 'add'),
  deleteRows: structureIcon('row', 'remove'),
  deleteCols: structureIcon('col', 'remove'),
  deleteCells: structureIcon('cell', 'remove'),

  // The outline family is built on the margin bracket the feature actually
  // draws, so it cannot be confused with insert/delete.
  outlineGroup: [
    ...sheet({ x: 8.4, y: 2.4, w: 13.2, h: 19.2, cols: 2, rows: 3 }),
    stroked('M6.4 3.4H3.4v8.6', PALETTE.accent, 'bold', { join: 'miter' }),
    stroked('M6.4 20.6H3.4V12', PALETTE.accent, 'bold', { join: 'miter' }),
    stroked('M1.4 12h4M3.4 10v4', PALETTE.accent, 'heavy'),
  ],
  outlineUngroup: [
    ...sheet({ x: 8.4, y: 2.4, w: 13.2, h: 19.2, cols: 2, rows: 3 }),
    stroked('M6.4 3.4H3.4v8.6', PALETTE.grid, 'bold', { join: 'miter' }),
    stroked('M6.4 20.6H3.4V12', PALETTE.grid, 'bold', { join: 'miter' }),
    stroked('M1.4 12h4', PALETTE.danger, 'heavy'),
  ],
  // Show detail expands the block; hide collapses it.
  outlineShow: [
    ...sheet({ x: 8.4, y: 2.4, w: 13.2, h: 19.2, cols: 2, rows: 4 }),
    ...arrow({
      from: [4, 11],
      to: [4, 2.8],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
    ...arrow({
      from: [4, 13],
      to: [4, 21.2],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
  ],
  outlineHide: [
    ...sheet({ x: 8.4, y: 6.4, w: 13.2, h: 11.2, cols: 2, rows: 2 }),
    ...arrow({
      from: [4, 2.8],
      to: [4, 11],
      color: PALETTE.danger,
      head: 'solid',
      weight: 'regular',
    }),
    ...arrow({
      from: [4, 21.2],
      to: [4, 13],
      color: PALETTE.danger,
      head: 'solid',
      weight: 'regular',
    }),
  ],

  // One column splits into two along the divider.
  textToColumns: [
    ...sheet({ x: 2.4, y: 3.4, w: 19.2, h: 17.2, cols: 2, rows: 3 }),
    stroked('M12 1.6v20.8', PALETTE.accent, 'bold', { dash: '2.8 2.2', cap: 'butt' }),
    ...arrow({
      from: [10.4, 12],
      to: [6.4, 12],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
    ...arrow({
      from: [13.6, 12],
      to: [17.6, 12],
      color: PALETTE.accent,
      head: 'solid',
      weight: 'regular',
    }),
  ],

  // Two identical rows, one of them struck out.
  removeDuplicates: [
    ...bars({ x: 3.4, y: 5, width: 12.6, widths: [1], gap: 0, color: PALETTE.ink }),
    ...bars({ x: 3.4, y: 10, width: 12.6, widths: [1], gap: 0, color: PALETTE.ink }),
    ...bars({ x: 3.4, y: 15, width: 12.6, widths: [1], gap: 0, color: PALETTE.grid }),
    ...bars({ x: 3.4, y: 20, width: 12.6, widths: [1], gap: 0, color: PALETTE.grid }),
    stroked('M2.4 12.4h14.6', PALETTE.danger, 'heavy'),
    ...badge({ glyph: 'cross', corner: 'tr' }),
  ],

  // Panes locked in place: the frozen bands sit above and left of the split.
  freeze: [
    ...sheet({
      x: 2.4,
      y: 2.4,
      w: 19.2,
      h: 19.2,
      cols: 3,
      rows: 3,
      bands: [
        { axis: 'row', index: 0, fill: PALETTE.infoSoft },
        { axis: 'col', index: 0, fill: PALETTE.infoSoft },
      ],
    }),
    stroked('M8.8 2.4v19.2M2.4 8.8h19.2', PALETTE.info, 'bold'),
  ],
} satisfies Record<string, IconDefinition>;
