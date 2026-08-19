/** Theme, page setup, print, zoom and window icons. */

import { circle, join, poly, rect, roundRect } from '../path.js';
import { badge, bars, filled, magnifier, outlined, place, sheet, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

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

  // A cog: one ring with the tooth gaps and the hub cut out by fill-rule.
  options: [
    {
      d: join(
        circle(12, 12, 9.4),
        circle(21.1, 12.0, 2.4),
        circle(18.43, 18.43, 2.4),
        circle(12.0, 21.1, 2.4),
        circle(5.57, 18.43, 2.4),
        circle(2.9, 12.0, 2.4),
        circle(5.57, 5.57, 2.4),
        circle(12.0, 2.9, 2.4),
        circle(18.43, 5.57, 2.4),
        circle(12, 12, 3.4),
      ),
      fill: PALETTE.info,
      fillRule: 'evenodd',
    },
  ],
} satisfies Record<string, IconDefinition>;
