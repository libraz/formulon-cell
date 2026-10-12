/** Theme, page setup, print, zoom and window icons. */

import { circle, ellipse, join, type Point, poly, rect, roundRect } from '../path.js';
import {
  arrow,
  badge,
  bars,
  doc,
  filled,
  magnifier,
  outlined,
  place,
  sheet,
  stroked,
} from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';
import { windowFrame } from './mac.js';

/** A themed window: title band, body text and two accent swatches. */
const themedWindow = (
  surface: string,
  band: string,
  text: string,
  swatches: readonly [string, string],
) => [
  outlined(rect(2.4, 3.4, 19.2, 17.2), surface, PALETTE.ink, 'regular'),
  filled(rect(2.4, 3.4, 19.2, 4.2), band),
  ...bars({
    x: 5,
    y: 10.4,
    width: 8.4,
    widths: [1, 0.75, 0.9],
    gap: 3.2,
    color: text,
    weight: 'regular',
  }),
  filled(rect(15.4, 9.4, 4.6, 4.2), swatches[0]),
  filled(rect(15.4, 14.6, 4.6, 4.2), swatches[1]),
];

/** A page with a folded corner, the base of the page-setup family. */
const page = (fold = 4.4) => [
  outlined(
    poly(
      [
        [4.4, 2.4],
        [15.2 + 4.4 - fold, 2.4],
        [19.6, 2.4 + fold],
        [19.6, 21.6],
        [4.4, 21.6],
      ],
      true,
    ),
    PALETTE.paper,
    PALETTE.ink,
    'regular',
  ),
  stroked(
    poly([
      [15.2 + 4.4 - fold, 2.4],
      [15.2 + 4.4 - fold, 2.4 + fold],
      [19.6, 2.4 + fold],
    ]),
    PALETTE.ink,
    'thin',
    { join: 'miter' },
  ),
];

/** Eight flat teeth around a ring; the hub is knocked out by the even-odd fill. */
const cog = (cx: number, cy: number, outer: number, inner: number, hub: number): string => {
  const at = (deg: number, r: number): Point => {
    const rad = (deg * Math.PI) / 180;
    return [cx + r * Math.cos(rad), cy + r * Math.sin(rad)];
  };
  const points: Point[] = [];
  for (let tooth = 0; tooth < 8; tooth++) {
    const mid = tooth * 45;
    points.push(at(mid - 15, inner), at(mid - 9, outer), at(mid + 9, outer), at(mid + 15, inner));
  }
  return join(poly(points, true), circle(cx, cy, hub));
};

/** A round plate in the bottom-right corner carrying a small paper-coloured mark. */
const cornerPlate = (tone: string, mark: readonly ReturnType<typeof stroked>[]) => [
  filled(circle(17.4, 17.4, 4.8), PALETTE.paper),
  filled(circle(17.4, 17.4, 4.2), tone),
  ...mark,
];

export const VIEW_ICONS = {
  themeLight: themedWindow(PALETTE.paper, PALETTE.accent, PALETTE.grid, [
    PALETTE.accentSoft,
    PALETTE.infoSoft,
  ]),
  themeDark: themedWindow(PALETTE.themeDarkSurface, PALETTE.themeDarkBand, PALETTE.themeDarkText, [
    PALETTE.info,
    PALETTE.alt,
  ]),
  themeContrast: themedWindow(PALETTE.themeContrastSurface, PALETTE.warn, PALETTE.paper, [
    PALETTE.warn,
    PALETTE.paper,
  ]),

  page: page(),
  pageSetup: [...page(), ...badge({ glyph: 'dots', corner: 'br', tone: PALETTE.info })],
  pageTheme: [
    outlined(rect(2.4, 3.4, 19.2, 17.2), PALETTE.paper, PALETTE.ink, 'regular'),
    filled(rect(2.4, 3.4, 19.2, 5.2), PALETTE.accent),
    ...bars({
      x: 5,
      y: 11.4,
      width: 8,
      widths: [1, 0.7],
      gap: 3.2,
      color: PALETTE.grid,
      weight: 'regular',
    }),
    filled(
      poly(
        [
          [17.4, 11.2],
          [20.6, 14.4],
          [17.4, 17.6],
          [14.2, 14.4],
        ],
        true,
      ),
      PALETTE.warn,
    ),
  ],
  printTitles: [
    ...sheet({
      x: 2.4,
      y: 3.4,
      w: 19.2,
      h: 17.2,
      cols: 3,
      rows: 4,
      bands: [{ axis: 'row', index: 0, fill: PALETTE.info }],
    }),
  ],
  printArea: [
    ...sheet({
      x: 2.4,
      y: 3.4,
      w: 19.2,
      h: 17.2,
      cols: 3,
      rows: 3,
      bands: [
        { axis: 'row', index: [0, 1], fill: PALETTE.accentSoft },
        { axis: 'col', index: 2, fill: PALETTE.paper },
      ],
    }),
    stroked(rect(2.4, 3.4, 12.8, 11.5), PALETTE.accent, 'bold', { join: 'miter', cap: 'butt' }),
  ],
  pageBreaks: [
    ...page(0),
    stroked('M4.4 12h15.2', PALETTE.info, 'heavy', { dash: '2.8 2.2' }),
    stroked('M12 2.4v19.2', PALETTE.info, 'heavy', { dash: '2.8 2.2' }),
  ],
  sheetBackground: [
    ...sheet({ x: 2.4, y: 3.4, w: 19.2, h: 17.2, cols: 3, rows: 3, surface: PALETTE.infoSoft }),
    filled(circle(17.2, 7.6, 1.7), PALETTE.warn),
    filled(
      poly(
        [
          [4.4, 19.4],
          [9.4, 12.4],
          [14.4, 19.4],
        ],
        true,
      ),
      PALETTE.accent,
    ),
  ],

  print: [
    outlined(rect(6.4, 2.4, 11.2, 5.4), PALETTE.paper, PALETTE.ink, 'regular'),
    outlined(roundRect(2.4, 7.8, 19.2, 8.4, 1.4), PALETTE.mute, PALETTE.ink, 'regular'),
    filled(circle(18.2, 11.4, 1.2), PALETTE.accent),
    outlined(rect(6.4, 14.4, 11.2, 7.2), PALETTE.paper, PALETTE.ink, 'regular'),
    ...bars({
      x: 8.4,
      y: 17.2,
      width: 7.2,
      widths: [1, 0.7],
      gap: 2.6,
      color: PALETTE.grid,
      weight: 'regular',
    }),
  ],

  zoom: [
    ...magnifier({ cx: 9.8, cy: 9.8, r: 7, handle: 6.4, handleColor: PALETTE.info }),
    stroked(join('M6.6 9.8h6.4', 'M9.8 6.6v6.4'), PALETTE.accent, 'bold'),
  ],

  // The pointer, at the size the rest of the set uses.
  objectSelect: [
    ...place(
      [
        outlined(
          poly(
            [
              [0, 0],
              [0, 15.4],
              [4.2, 11.6],
              [6.8, 17.4],
              [9.4, 16.2],
              [6.8, 10.6],
              [12.2, 10.2],
            ],
            true,
          ),
          PALETTE.paper,
          PALETTE.ink,
          'bold',
        ),
      ],
      { box: [0, 0, 12.2, 17.4], size: 19.4 },
    ),
  ],
  selectionPane: [
    outlined(rect(2.4, 3.4, 19.2, 17.2), PALETTE.paper, PALETTE.ink, 'regular'),
    stroked('M12 3.4v17.2', PALETTE.grid, 'thin'),
    ...bars({
      x: 4.6,
      y: 7.4,
      width: 5.2,
      widths: [1, 1, 1],
      gap: 3.4,
      color: PALETTE.grid,
      weight: 'regular',
    }),
    ...bars({
      x: 14.2,
      y: 7.4,
      width: 5.2,
      widths: [1, 1, 1],
      gap: 3.4,
      color: PALETTE.info,
      weight: 'regular',
    }),
  ],

  options: [{ d: cog(12, 12, 9.6, 7.2, 3.2), fill: PALETTE.info, fillRule: 'evenodd' }],

  help: [
    outlined(circle(12, 12, 9.2), PALETTE.paper, PALETTE.info, 'bold'),
    stroked(
      'M9.3 9.5c0-1.7 1.2-2.8 2.8-2.8s2.8 1 2.8 2.5c0 2-2.8 2.1-2.8 4.3',
      PALETTE.info,
      'bold',
    ),
    filled(circle(12.1, 16.8, 1.05), PALETTE.info),
  ],

  // Row and column headers both numbered, one cell addressed by them.
  referenceStyle: [
    ...sheet({
      x: 2.4,
      y: 3.4,
      w: 19.2,
      h: 17.2,
      cols: 4,
      rows: 4,
      bands: [
        { axis: 'row', index: 0, fill: PALETTE.infoSoft },
        { axis: 'col', index: 0, fill: PALETTE.infoSoft },
      ],
    }),
    stroked(rect(12, 12, 4.8, 4.3), PALETTE.info, 'bold', { join: 'miter', cap: 'butt' }),
  ],

  zeroValues: [
    ...sheet({ x: 2.4, y: 3.4, w: 19.2, h: 17.2, cols: 1, rows: 1 }),
    stroked(ellipse(12, 12, 3.4, 4.9), PALETTE.ink, 'heavy'),
  ],

  rightToLeft: [
    stroked(join('M7.2 4.8h13.2', 'M11.4 8.8h9', 'M7.2 12.8h13.2'), PALETTE.ink, 'bold', {
      cap: 'butt',
    }),
    ...arrow({ from: [20.4, 18.6], to: [3.6, 18.6], color: PALETTE.info, head: 'solid' }),
  ],

  bringForward: [
    outlined(roundRect(2.6, 2.6, 11.6, 11.6, 1.2), PALETTE.paper, PALETTE.grid, 'regular'),
    outlined(roundRect(9.8, 9.8, 11.6, 11.6, 1.2), PALETTE.infoSoft, PALETTE.info, 'regular'),
  ],

  whatIf: [
    ...sheet({ x: 2.4, y: 3.4, w: 15.2, h: 13.6, cols: 3, rows: 3 }),
    ...cornerPlate(PALETTE.info, [
      stroked(
        'M15.8 16.4c0-1 .7-1.6 1.6-1.6s1.6.6 1.6 1.4c0 1.1-1.6 1.2-1.6 2.3',
        PALETTE.paper,
        'regular',
      ),
      filled(circle(17.4, 20.1, 0.6), PALETTE.paper),
    ]),
  ],

  showChanges: [
    ...doc({ x: 4, y: 2.4, w: 13, h: 16.4, fold: 3.6, lines: 3 }),
    ...cornerPlate(PALETTE.info, [
      stroked(join('M17.4 17.4v-2.4', 'M17.4 17.4l1.8 1.2'), PALETTE.paper, 'regular'),
    ]),
  ],

  customViews: [
    ...windowFrame({ x: 2.4, y: 3.4, w: 16.6, h: 14.6 }),
    ...badge({ glyph: 'star', corner: 'br' }),
  ],
} satisfies Record<string, IconDefinition>;
