/** Table, pivot-table and cell-style icons. */

import { hLine, join, vLine } from '../path.js';
import { badge, brush, headed, sheet, stroked } from '../primitives.js';
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

/** The grid the style and format commands act on: four columns by three rows. */
const STYLED = {
  x: 2.4,
  y: 3.4,
  w: 19.2,
  h: 14.4,
  cols: 4,
  rows: 3,
  rules: PALETTE.grid,
} as const;

/** The brush that lays a style down, crossing the grid's lower-right corner. */
const styleBrush = brush({ bristle: PALETTE.info, cx: 16.4, cy: 15.6, size: 13.6, rotate: 40 });

/** A four-by-four table with the region a Table Design option switches on picked out. */
const option = (bands: NonNullable<Parameters<typeof sheet>[0]>['bands']) =>
  sheet({ x: 2.4, y: 3.4, w: 19.2, h: 17.2, cols: 4, rows: 4, bands });

export const TABLE_ICONS = {
  table: table(FULL, PALETTE.accent),
  tableStyle: [
    ...sheet({
      ...STYLED,
      cells: [{ cols: [1, 2], rows: [1, 2], fill: PALETTE.infoMid, frame: PALETTE.info }],
    }),
    ...styleBrush,
  ],

  // Table Design options: the header, total, stripe or edge band each one adds.
  tableHeaderRow: option([{ axis: 'row', index: 0, fill: PALETTE.accent }]),
  tableTotalRow: option([{ axis: 'row', index: 3, fill: PALETTE.accent }]),
  tableBandedRows: option([
    { axis: 'row', index: 1, fill: PALETTE.infoSoft },
    { axis: 'row', index: 3, fill: PALETTE.infoSoft },
  ]),
  tableBandedColumns: option([
    { axis: 'col', index: 1, fill: PALETTE.infoSoft },
    { axis: 'col', index: 3, fill: PALETTE.infoSoft },
  ]),
  tableFirstColumn: option([{ axis: 'col', index: 0, fill: PALETTE.accent }]),
  tableLastColumn: option([{ axis: 'col', index: 3, fill: PALETTE.accent }]),

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

  // Format cells: one cell in focus, with the dimension it is sized by.
  formatCells: [
    ...sheet({
      ...STYLED,
      cols: 3,
      rows: 3,
      y: 5.6,
      h: 16,
      cells: [{ cols: [1, 1], rows: [1, 1], fill: PALETTE.infoMid, frame: PALETTE.info }],
    }),
    stroked(join(hLine(9, 2.8, 6), vLine(9, 1.6, 2.4), vLine(15, 1.6, 2.4)), PALETTE.info, 'thin', {
      cap: 'butt',
    }),
  ],
  // Cells a rule has flagged, picked out in the rule's colour.
  conditional: [
    ...sheet({
      ...STYLED,
      y: 4.8,
      cells: [
        { cols: [0, 1], rows: [0, 0], fill: PALETTE.dangerMid, frame: PALETTE.danger },
        { cols: [0, 0], rows: [1, 1], fill: PALETTE.infoMid, frame: PALETTE.info },
        { cols: [0, 1], rows: [2, 2], fill: PALETTE.dangerMid, frame: PALETTE.danger },
      ],
    }),
  ],
  // A style brushed onto a block of cells.
  cellStyles: [
    ...sheet({
      ...STYLED,
      cells: [{ cols: [0, 2], rows: [0, 1], fill: PALETTE.infoMid, frame: PALETTE.info }],
    }),
    ...styleBrush,
  ],
} satisfies Record<string, IconDefinition>;
