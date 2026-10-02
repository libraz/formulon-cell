/** Mac ribbon additions that are not part of the shared Excel glyph family. */

import { circle, ellipse, join, poly, rect, roundRect } from '../path.js';
import { arcArrow, filled, outlined, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

const MAC_ICON_KEYS = [
  'form',
  'icons',
  'threeD',
  'smartArt',
  'checkBox',
  'chartHierarchy',
  'chartStatistical',
  'chartWaterfall',
  'chartCombo',
  'chartMap',
  'sparkline',
  'slicer',
  'timeline',
  'textBox',
  'headerFooter',
  'wordArt',
  'object',
  'symbol',
  'lasso',
  'pencil',
  'penBlack',
  'penRed',
  'highlighter',
  'plus',
  'trackpad',
  'font',
  'python',
  'data',
  'refresh',
  'count',
  'window',
  'arrange',
  'split',
  'hide',
  'show',
  'code',
] as const;

export type MacIconKey = (typeof MAC_ICON_KEYS)[number];

/** The Mac model's icon vocabulary; useful to keep the audit and glyph set in lockstep. */
export { MAC_ICON_KEYS };

const chartAxes = () => stroked('M3.2 3.2v17.6h17.6', PALETTE.ink, 'regular', { cap: 'square' });

const chartBar = (x: number, y: number, w: number, h: number, fill: string) =>
  filled(roundRect(x, y, w, h, 0.7), fill);

const node = (x: number, y: number, fill: string, radius = 2.4) =>
  outlined(circle(x, y, radius), PALETTE.paper, fill, 'bold');

const windowFrame = (options: { x?: number; y?: number; w?: number; h?: number } = {}) => {
  const x = options.x ?? 2.4;
  const y = options.y ?? 3.4;
  const w = options.w ?? 19.2;
  const h = options.h ?? 17.2;
  return [
    outlined(roundRect(x, y, w, h, 1.4), PALETTE.paper, PALETTE.ink, 'regular'),
    filled(roundRect(x, y, w, 4, 1.4), PALETTE.infoSoft),
    filled(circle(x + 2.4, y + 2, 0.75), PALETTE.danger),
    filled(circle(x + 4.8, y + 2, 0.75), PALETTE.warnDeep),
    filled(circle(x + 7.2, y + 2, 0.75), PALETTE.accent),
  ];
};

const eye = (slash = false) => [
  stroked(
    'M2.4 12c3.1-4.2 6.3-6.4 9.6-6.4s6.5 2.2 9.6 6.4c-3.1 4.2-6.3 6.4-9.6 6.4S5.5 16.2 2.4 12z',
    PALETTE.ink,
    'regular',
  ),
  filled(circle(12, 12, 3), PALETTE.infoSoft),
  filled(circle(12, 12, 1.4), PALETTE.infoDeep),
  ...(slash ? [stroked('M3.2 3.2 20.8 20.8', PALETTE.danger, 'heavy')] : []),
];

const codeMark = (color: string) => [
  stroked(
    poly([
      [8, 5.6],
      [3.4, 12],
      [8, 18.4],
    ]),
    color,
    'heavy',
  ),
  stroked(
    poly([
      [16, 5.6],
      [20.6, 12],
      [16, 18.4],
    ]),
    color,
    'heavy',
  ),
  stroked('M14.4 4.6 9.6 19.4', PALETTE.accent, 'bold'),
];

const penMark = (ink: string, body: string) => [
  filled(
    poly(
      [
        [4.2, 18.8],
        [5.8, 20.4],
        [19.8, 6.4],
        [18.2, 4.8],
      ],
      true,
    ),
    body,
  ),
  filled(
    poly(
      [
        [4.2, 18.8],
        [5.8, 20.4],
        [3.2, 21.6],
      ],
      true,
    ),
    PALETTE.tan,
  ),
  stroked(
    poly([
      [3.2, 21.6],
      [4.2, 18.8],
      [5.8, 20.4],
      [3.2, 21.6],
    ]),
    PALETTE.ink,
    'regular',
  ),
  filled(
    poly(
      [
        [18.2, 4.8],
        [19.8, 6.4],
        [21.2, 5],
        [19.6, 3.6],
      ],
      true,
    ),
    ink,
  ),
  stroked('M7.4 18.6 17.8 8.2', PALETTE.paper, 'thin'),
];

export const MAC_ICONS: Record<MacIconKey, IconDefinition> = {
  /** A form card with mixed controls and labels. */
  form: [
    outlined(roundRect(2.4, 2.4, 19.2, 19.2, 1.4), PALETTE.paper, PALETTE.ink, 'regular'),
    filled(roundRect(2.4, 2.4, 19.2, 4, 1.4), PALETTE.infoSoft),
    filled(circle(5, 4.4, 0.75), PALETTE.info),
    stroked(
      join('M6.8 9.2h10.8', 'M6.8 12.4h8', 'M6.8 15.6h10.8', 'M6.8 18.8h6.4'),
      PALETTE.grid,
      'regular',
    ),
    outlined(roundRect(3.8, 8, 1.8, 1.8, 0.35), PALETTE.paper, PALETTE.accent, 'thin'),
    stroked('M4.2 8.9 4.7 9.4 5.4 8.5', PALETTE.accent, 'bold'),
    outlined(circle(4.7, 12.4, 0.9), PALETTE.paper, PALETTE.info, 'thin'),
    outlined(roundRect(3.8, 14.7, 1.8, 1.8, 0.35), PALETTE.paper, PALETTE.grid, 'thin'),
  ],

  /** A gallery of four small pictograms. */
  icons: [
    outlined(roundRect(2.4, 2.4, 19.2, 19.2, 1.4), PALETTE.paper, PALETTE.ink, 'regular'),
    filled(circle(7.3, 7.3, 2.6), PALETTE.infoSoft),
    filled(
      poly(
        [
          [7.3, 4.1],
          [8.1, 6.5],
          [10.5, 7.3],
          [8.1, 8.1],
          [7.3, 10.5],
          [6.5, 8.1],
          [4.1, 7.3],
          [6.5, 6.5],
        ],
        true,
      ),
      PALETTE.info,
    ),
    filled(circle(16.7, 7.3, 2.6), PALETTE.warnSoft),
    filled(
      poly(
        [
          [16.7, 4.1],
          [17.5, 6.5],
          [19.9, 7.3],
          [17.5, 8.1],
          [16.7, 10.5],
          [15.9, 8.1],
          [13.5, 7.3],
          [15.9, 6.5],
        ],
        true,
      ),
      PALETTE.warnDeep,
    ),
    outlined(roundRect(4.2, 13.8, 5.9, 4.8, 0.8), PALETTE.accentSoft, PALETTE.accent, 'thin'),
    stroked('M5.6 16.2h3.1', PALETTE.accentDeep, 'regular'),
    outlined(roundRect(13.9, 13.8, 5.9, 4.8, 0.8), PALETTE.violetSoft, PALETTE.violet, 'thin'),
    stroked('M15.3 16.2h3.1', PALETTE.violet, 'regular'),
  ],

  /** A solid isometric cube for 3D content. */
  threeD: [
    filled(
      poly(
        [
          [12, 2.8],
          [21, 7.6],
          [12, 12.6],
          [3, 7.6],
        ],
        true,
      ),
      PALETTE.infoSoft,
    ),
    filled(
      poly(
        [
          [3, 7.6],
          [12, 12.6],
          [12, 21.2],
          [3, 16.2],
        ],
        true,
      ),
      PALETTE.accentSoft,
    ),
    filled(
      poly(
        [
          [12, 12.6],
          [21, 7.6],
          [21, 16.2],
          [12, 21.2],
        ],
        true,
      ),
      PALETTE.warnSoft,
    ),
    stroked(
      join('M12 2.8 21 7.6v8.6L12 21.2 3 16.2V7.6z', 'M12 12.6v8.6'),
      PALETTE.ink,
      'regular',
      { join: 'miter' },
    ),
  ],

  /** A connected process diagram with a central branch. */
  smartArt: [
    stroked(
      join('M12 7.2v3.2', 'M12 10.4H6.6v3', 'M12 10.4h5.4v3', 'M6.6 17.4h10.8'),
      PALETTE.info,
      'bold',
    ),
    outlined(roundRect(8.2, 2.4, 7.6, 4.8, 1.2), PALETTE.infoSoft, PALETTE.info, 'regular'),
    outlined(roundRect(3, 13.4, 7.2, 5.4, 1.2), PALETTE.accentSoft, PALETTE.accent, 'regular'),
    outlined(roundRect(13.8, 13.4, 7.2, 5.4, 1.2), PALETTE.warnSoft, PALETTE.warnDeep, 'regular'),
    stroked('M5 16.1h3.2M15.8 16.1H19', PALETTE.grid, 'thin'),
  ],

  /** A worksheet with a checked form control. */
  checkBox: [
    ...sheet({ x: 2.4, y: 3.4, w: 19.2, h: 17.2, cols: 1, rows: 4 }),
    outlined(roundRect(4.5, 5.5, 4.2, 4.2, 0.55), PALETTE.accentSoft, PALETTE.accent, 'regular'),
    stroked('M5.3 7.6 6.4 8.7 8 6.8', PALETTE.accentDeep, 'heavy'),
    stroked(join('M11.4 7.6h7', 'M11.4 12.2h5.4', 'M11.4 16.8h7'), PALETTE.grid, 'regular'),
  ],

  /** A hierarchy chart with three levels of connected nodes. */
  chartHierarchy: [
    chartAxes(),
    stroked(
      join('M12 5.4v3.2', 'M12 8.6H7.2v3.4', 'M12 8.6h4.8v3.4', 'M7.2 15.4v2.8', 'M16.8 15.4v2.8'),
      PALETTE.info,
      'regular',
    ),
    node(12, 4.6, PALETTE.accent, 2),
    node(7.2, 13.4, PALETTE.info, 2),
    node(16.8, 13.4, PALETTE.warnDeep, 2),
    filled(rect(5.2, 18.2, 4, 1.3), PALETTE.infoSoft),
    filled(rect(14.8, 18.2, 4, 1.3), PALETTE.warnSoft),
  ],

  /** A box-and-whisker statistical plot. */
  chartStatistical: [
    chartAxes(),
    stroked(
      join('M6.4 7.2v10.4', 'M5.2 7.2h2.4', 'M5.2 17.6h2.4', 'M4.2 11h4.4'),
      PALETTE.info,
      'regular',
    ),
    outlined(rect(10, 8.2, 4, 7), PALETTE.accentSoft, PALETTE.accent, 'regular'),
    stroked(
      join('M10 11.8h4', 'M12 6.4v1.8', 'M12 15.2v3.2', 'M10.8 6.4h2.4', 'M10.8 18.4h2.4'),
      PALETTE.accentDeep,
      'regular',
    ),
    stroked(
      join('M17.4 5.2v12.6', 'M16.2 5.2h2.4', 'M16.2 17.8h2.4', 'M15.2 10.6h4.4'),
      PALETTE.warnDeep,
      'regular',
    ),
  ],

  /** Incremental bars connected by a running-total bridge. */
  chartWaterfall: [
    chartAxes(),
    chartBar(5, 14.4, 3.2, 4.6, PALETTE.info),
    chartBar(9.4, 9.6, 3.2, 4.8, PALETTE.accent),
    chartBar(13.8, 12.8, 3.2, 3.2, PALETTE.dangerSoft),
    chartBar(18.2, 5.6, 3.2, 11.6, PALETTE.warnDeep),
    stroked(join('M8.2 14.4h1.2', 'M12.6 9.6h1.2', 'M17 12.8h1.2'), PALETTE.grid, 'thin', {
      dash: '1.6 1.4',
    }),
  ],

  /** Bars and a line share one set of axes. */
  chartCombo: [
    chartAxes(),
    chartBar(5.2, 12.4, 3.2, 7.6, PALETTE.infoSoft),
    chartBar(10.4, 8.2, 3.2, 11.8, PALETTE.accentSoft),
    chartBar(15.6, 14.6, 3.2, 5.4, PALETTE.warnSoft),
    stroked('M6.8 10.4 12 5.2 17.2 9.4 20.4 4.6', PALETTE.violet, 'bold'),
    filled(
      join(
        circle(6.8, 10.4, 1.1),
        circle(12, 5.2, 1.1),
        circle(17.2, 9.4, 1.1),
        circle(20.4, 4.6, 1.1),
      ),
      PALETTE.violet,
    ),
  ],

  /** A map panel with a highlighted region and location pin. */
  chartMap: [
    outlined(roundRect(2.4, 3.2, 19.2, 17.6, 1.2), PALETTE.infoSoft, PALETTE.info, 'regular'),
    stroked(
      join('M8.8 3.2v17.6', 'M15.2 3.2v17.6', 'M2.4 9.1h19.2', 'M2.4 14.9h19.2'),
      PALETTE.gridLight,
      'hairline',
      { cap: 'butt' },
    ),
    filled(
      poly(
        [
          [4.2, 6.6],
          [7.4, 5.4],
          [9.4, 8],
          [8, 11.4],
          [5, 10.8],
        ],
        true,
      ),
      PALETTE.accentSoft,
    ),
    filled(
      poly(
        [
          [16.2, 12],
          [20.2, 10.8],
          [20.8, 15.4],
          [17.6, 17.8],
          [15.4, 15.2],
        ],
        true,
      ),
      PALETTE.warnSoft,
    ),
    filled(circle(13.2, 9.4, 2.5), PALETTE.danger),
    filled(circle(13.2, 9.4, 0.85), PALETTE.paper),
    stroked('M13.2 11.8v3.4', PALETTE.danger, 'regular'),
  ],

  /** A compact sparkline without chart axes. */
  sparkline: [
    stroked('M2.6 17.8 6.2 15.6 9.4 16.8 12.6 9.2 16.2 11.2 21.4 4.8', PALETTE.info, 'bold'),
    filled(
      join(
        circle(2.6, 17.8, 1),
        circle(6.2, 15.6, 1),
        circle(9.4, 16.8, 1),
        circle(12.6, 9.2, 1),
        circle(16.2, 11.2, 1),
        circle(21.4, 4.8, 1),
      ),
      PALETTE.accent,
    ),
  ],

  /** A filter panel with selectable slicer chips. */
  slicer: [
    outlined(roundRect(2.4, 2.4, 19.2, 19.2, 1.2), PALETTE.paper, PALETTE.ink, 'regular'),
    filled(roundRect(2.4, 2.4, 19.2, 3.8, 1.2), PALETTE.infoSoft),
    stroked('M4.8 8.8h14.4', PALETTE.grid, 'regular'),
    outlined(roundRect(4.4, 10.4, 6.2, 3.6, 0.8), PALETTE.accentSoft, PALETTE.accent, 'thin'),
    outlined(roundRect(13.4, 10.4, 6.2, 3.6, 0.8), PALETTE.infoSoft, PALETTE.info, 'thin'),
    outlined(roundRect(4.4, 16, 6.2, 3.6, 0.8), PALETTE.paper, PALETTE.grid, 'thin'),
    outlined(roundRect(13.4, 16, 6.2, 3.6, 0.8), PALETTE.warnSoft, PALETTE.warnDeep, 'thin'),
    stroked('M6 12.2h3M15 12.2h3M6 17.8h3M15 17.8h3', PALETTE.grid, 'hairline'),
  ],

  /** A time axis with events and a highlighted interval. */
  timeline: [
    stroked('M3.2 13.2h17.6', PALETTE.ink, 'regular'),
    filled(roundRect(5.2, 6.2, 6.4, 2.8, 1.2), PALETTE.infoSoft),
    filled(roundRect(14.2, 16.2, 5.2, 2.8, 1.2), PALETTE.accentSoft),
    filled(circle(5.2, 13.2, 2), PALETTE.info),
    filled(circle(12, 13.2, 2), PALETTE.accent),
    filled(circle(19, 13.2, 2), PALETTE.warnDeep),
    stroked('M5.2 10v3.2M12 15.2v-2M19 13.2v3', PALETTE.grid, 'thin'),
  ],

  /** A text box with editable lines. */
  textBox: [
    outlined(roundRect(2.4, 4, 19.2, 16, 1.4), PALETTE.paper, PALETTE.info, 'bold'),
    stroked(join('M6 9.2h12', 'M6 12.4h9.2', 'M6 15.6h12'), PALETTE.ink, 'regular'),
    stroked('M6 18.8h5.6', PALETTE.info, 'regular'),
  ],

  /** A page showing distinct header and footer bands. */
  headerFooter: [
    ...sheet({ x: 4, y: 2.4, w: 16, h: 19.2, cols: 2, rows: 5 }),
    filled(roundRect(4, 2.4, 16, 3.2, 0.7), PALETTE.infoSoft),
    filled(roundRect(4, 18.4, 16, 3.2, 0.7), PALETTE.accentSoft),
    stroked('M6 4h5M15 20h3', PALETTE.infoDeep, 'regular'),
  ],

  /** A decorative slanted word-art ribbon represented by outlined strokes. */
  wordArt: [
    outlined(
      poly(
        [
          [2.6, 7],
          [19.6, 3.2],
          [21.4, 16.8],
          [4.4, 20.6],
        ],
        true,
      ),
      PALETTE.warnSoft,
      PALETTE.warnDeep,
      'regular',
    ),
    stroked(
      join('M6.2 14.4 8.2 9.2l2.2 4.2 2.4-5.2 2.2 4.2 2.1-4.6', 'M6.6 17h10.8'),
      PALETTE.violet,
      'bold',
    ),
  ],

  /** An embedded object card with an attachment marker. */
  object: [
    ...windowFrame({ x: 2.4, y: 3.2, w: 18.2, h: 16.4 }),
    outlined(roundRect(7, 8.2, 10, 7.2, 1), PALETTE.paper, PALETTE.grid, 'regular'),
    stroked('M9.2 11.2h5.6M9.2 13.6h3.8', PALETTE.grid, 'regular'),
    filled(circle(19.2, 18.4, 2.8), PALETTE.accent),
    stroked(join('M17.8 18.4h2.8', 'M19.2 17v2.8'), PALETTE.paper, 'heavy'),
  ],

  /** A symbol palette: star, ring and diamond marks. */
  symbol: [
    filled(
      poly(
        [
          [7, 3.2],
          [8.5, 7.2],
          [12.8, 7.4],
          [9.4, 10],
          [10.4, 14.2],
          [7, 11.8],
          [3.6, 14.2],
          [4.6, 10],
          [1.2, 7.4],
          [5.5, 7.2],
        ],
        true,
      ),
      PALETTE.info,
    ),
    stroked(circle(16.6, 8.6, 4.2), PALETTE.accent, 'bold'),
    filled(
      poly(
        [
          [16.6, 15],
          [20.6, 19],
          [16.6, 23],
          [12.6, 19],
        ],
        true,
      ),
      PALETTE.warnSoft,
    ),
    stroked(
      poly([
        [16.6, 15],
        [20.6, 19],
        [16.6, 23],
        [12.6, 19],
        [16.6, 15],
      ]),
      PALETTE.warnDeep,
      'regular',
    ),
  ],

  /** A lasso loop around a selected point and a pointer. */
  lasso: [
    stroked(ellipse(11.4, 11.4, 7.8, 5.8), PALETTE.info, 'bold', { dash: '2.2 1.8' }),
    stroked('M17.8 16.8 21 21.2', PALETTE.ink, 'bold'),
    filled(
      poly(
        [
          [14.4, 14.6],
          [18.8, 21.8],
          [19, 17.8],
          [22, 16.8],
        ],
        true,
      ),
      PALETTE.paper,
    ),
    stroked(
      poly([
        [14.4, 14.6],
        [18.8, 21.8],
        [19, 17.8],
        [22, 16.8],
        [14.4, 14.6],
      ]),
      PALETTE.ink,
      'regular',
    ),
  ],

  /** A graphite pencil with a sharpened tip. */
  pencil: [
    filled(
      poly(
        [
          [4, 18.8],
          [5.6, 20.4],
          [19.8, 6.2],
          [18.2, 4.6],
        ],
        true,
      ),
      PALETTE.grid,
    ),
    filled(
      poly(
        [
          [4, 18.8],
          [5.6, 20.4],
          [3.2, 21.6],
        ],
        true,
      ),
      PALETTE.tan,
    ),
    stroked(
      poly([
        [3.2, 21.6],
        [4, 18.8],
        [5.6, 20.4],
        [3.2, 21.6],
      ]),
      PALETTE.ink,
      'regular',
    ),
    filled(
      poly(
        [
          [18.2, 4.6],
          [19.8, 6.2],
          [21.2, 4.8],
          [19.6, 3.4],
        ],
        true,
      ),
      PALETTE.info,
    ),
    stroked('M7.4 18.6 17.8 8.2', PALETTE.paper, 'thin'),
  ],

  penBlack: penMark(PALETTE.ink, PALETTE.grid),
  penRed: penMark(PALETTE.danger, PALETTE.dangerSoft),

  /** A broad yellow marker with a translucent chisel tip. */
  highlighter: [
    filled(
      poly(
        [
          [4.2, 18.8],
          [6.4, 21],
          [20.4, 7],
          [18.2, 4.8],
        ],
        true,
      ),
      PALETTE.warn,
    ),
    outlined(
      poly(
        [
          [4.2, 18.8],
          [6.4, 21],
          [20.4, 7],
          [18.2, 4.8],
          [4.2, 18.8],
        ],
        true,
      ),
      PALETTE.warnDeep,
      PALETTE.warnDeep,
      'thin',
    ),
    stroked('M8.4 18.8 18.2 9', PALETTE.warnSoft, 'regular'),
    stroked('M3.2 21.4h4.6', PALETTE.warnDeep, 'bold'),
  ],

  plus: [
    outlined(circle(12, 12, 9.1), PALETTE.paper, PALETTE.accent, 'thin'),
    stroked(join('M4.2 12h15.6', 'M12 4.2v15.6'), PALETTE.accent, 'heavy'),
  ],

  /** A trackpad with touch points and a gesture arc. */
  trackpad: [
    outlined(roundRect(2.4, 3.2, 19.2, 17.6, 2.2), PALETTE.mute, PALETTE.ink, 'regular'),
    stroked('M4.6 16.8h15', PALETTE.grid, 'thin'),
    filled(circle(8, 11, 1.8), PALETTE.info),
    filled(circle(12, 9, 1.8), PALETTE.accent),
    filled(circle(16, 11, 1.8), PALETTE.warnDeep),
    stroked('M6 7.2c2.6-2.4 5.2-3.2 8.2-2.2', PALETTE.info, 'regular'),
  ],

  /** A geometric typeface mark, built from stems and a crossbar. */
  font: [
    stroked(join('M4.2 19.6 9.6 4.4h4.8l5.4 15.2', 'M7.4 13.6h9.2'), PALETTE.ink, 'heavy'),
    stroked('M10.4 4.4h3.2', PALETTE.info, 'regular'),
    stroked('M5.8 19.6h3.2M15 19.6h3.2', PALETTE.accent, 'regular'),
  ],

  /** Interlocking rounded snakes evoke Python without relying on text glyphs. */
  python: [
    stroked(
      'M5.2 11.8V8.2c0-2.2 1.8-4 4-4h3.2v3.2H9.6c-.7 0-1.2.5-1.2 1.2v3.2h5.4c2.2 0 4 1.8 4 4v.8',
      PALETTE.info,
      'heavy',
    ),
    stroked(
      'M18.8 12.2v3.6c0 2.2-1.8 4-4 4h-3.2v-3.2h2.8c.7 0 1.2-.5 1.2-1.2v-3.2h-5.4c-2.2 0-4-1.8-4-4v-.8',
      PALETTE.warnDeep,
      'heavy',
    ),
    filled(circle(10, 6.4, 0.8), PALETTE.info),
    filled(circle(14, 17.6, 0.8), PALETTE.warnDeep),
  ],

  /** A small database cylinder over a data grid. */
  data: [
    filled(ellipse(12, 5.8, 7.4, 2.7), PALETTE.infoSoft),
    filled(rect(4.6, 5.8, 14.8, 10.8), PALETTE.infoSoft),
    filled(ellipse(12, 16.6, 7.4, 2.7), PALETTE.infoSoft),
    stroked(
      join('M4.6 5.8v10.8', 'M19.4 5.8v10.8', 'M4.6 9.4c0 1.5 3.3 2.7 7.4 2.7s7.4-1.2 7.4-2.7'),
      PALETTE.infoDeep,
      'regular',
    ),
    stroked('M7.2 19.2h9.6', PALETTE.accent, 'bold'),
  ],

  refresh: [
    ...arcArrow({
      from: [4.8, 16.8],
      to: [19.2, 7.2],
      radius: 9.6,
      sweep: 1,
      headDir: [0.9, -0.4],
      color: PALETTE.info,
      weight: 'bold',
      head: 'solid',
    }),
    stroked('M3.4 7.8h5.2M3.4 7.8v5.2', PALETTE.accent, 'bold'),
  ],

  /** A counting mark made from grouped tally strokes. */
  count: [
    outlined(roundRect(2.8, 3.2, 18.4, 17.6, 1.2), PALETTE.paper, PALETTE.ink, 'regular'),
    stroked(
      join('M6 8.2v4.2', 'M9 8.2v4.2', 'M12 8.2v4.2', 'M15 8.2v4.2', 'M5.2 12.8 15.8 7.6'),
      PALETTE.info,
      'bold',
    ),
    stroked('M6 16.4h11.8', PALETTE.accent, 'regular'),
    filled(circle(18.4, 16.4, 1.2), PALETTE.warnDeep),
  ],

  window: windowFrame(),

  /** Two overlapping windows with direction arrows indicate arrangement. */
  arrange: [
    ...windowFrame({ x: 2.4, y: 3.2, w: 13.8, h: 13.2 }),
    outlined(roundRect(9.8, 8.2, 11.8, 12.4, 1.2), PALETTE.paper, PALETTE.accent, 'regular'),
    stroked(join('M4.8 12.8h6.2', 'M8.8 10.2 11.4 12.8 8.8 15.4'), PALETTE.info, 'bold'),
    stroked(join('M14 17h4.8', 'M16.8 14.4 19.4 17 16.8 19.6'), PALETTE.accent, 'bold'),
  ],

  /** Window panes separated by a visible cross-shaped splitter. */
  split: [
    ...windowFrame(),
    filled(rect(11.2, 7.4, 1.6, 12.2), PALETTE.info),
    filled(rect(4.2, 12.4, 15.2, 1.6), PALETTE.info),
    stroked('M12 11.2v3.8M10.2 13.2h3.6', PALETTE.paper, 'thin'),
  ],

  hide: eye(true),
  show: eye(false),

  code: codeMark(PALETTE.info),
};
