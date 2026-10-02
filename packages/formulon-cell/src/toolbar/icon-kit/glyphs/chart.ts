/** Chart-type icons, all sharing one plot frame. */

import { circle, join, line, poly, rect, sector } from '../path.js';
import { badge, filled, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

/** Plot area: the axes every chart icon is drawn against. */
const PLOT = { left: 3.6, right: 21.4, top: 2.6, bottom: 19.4 } as const;

const axes = () =>
  stroked(
    `M${PLOT.left} ${PLOT.top}v${PLOT.bottom - PLOT.top}h${PLOT.right - PLOT.left}`,
    PALETTE.ink,
    'regular',
    { cap: 'square' },
  );

/** A column rising from the baseline. */
const column = (x: number, width: number, height: number, fill: string) =>
  filled(rect(x, PLOT.bottom - height, width, height), fill);

const PIE_CENTER = 12;
const PIE_RADIUS = 8.8;
const piePoint = (radians: number): [number, number] => [
  PIE_CENTER + PIE_RADIUS * Math.cos(radians),
  PIE_CENTER + PIE_RADIUS * Math.sin(radians),
];

export const CHART_ICONS = {
  chart: [
    axes(),
    column(6, 3.8, 6.6, PALETTE.info),
    column(11.1, 3.8, 11, PALETTE.accent),
    column(16.2, 3.8, 14.6, PALETTE.alt),
  ],
  chartColumn: [
    axes(),
    column(6, 3.8, 6.6, PALETTE.info),
    column(11.1, 3.8, 11, PALETTE.accent),
    column(16.2, 3.8, 14.6, PALETTE.alt),
  ],
  chartRecommended: [
    axes(),
    column(6, 3.4, 6, PALETTE.info),
    column(10.4, 3.4, 10, PALETTE.accent),
    column(14.8, 3.4, 13.4, PALETTE.alt),
    ...badge({ glyph: 'star', corner: 'tr', tone: PALETTE.warn, onTone: PALETTE.warnDeep }),
  ],
  chartBar: [
    axes(),
    filled(rect(PLOT.left + 0.8, 4.2, 13.6, 3.6), PALETTE.info),
    filled(rect(PLOT.left + 0.8, 9.2, 9, 3.6), PALETTE.accent),
    filled(rect(PLOT.left + 0.8, 14.2, 15.8, 3.6), PALETTE.alt),
  ],
  chartLine: [
    axes(),
    stroked('M6 15.4 10.4 11 14.4 13.4 19.6 5.8', PALETTE.info, 'bold'),
    filled(
      join(
        circle(6, 15.4, 1.5),
        circle(10.4, 11, 1.5),
        circle(14.4, 13.4, 1.5),
        circle(19.6, 5.8, 1.5),
      ),
      PALETTE.info,
    ),
  ],
  chartArea: [
    axes(),
    filled(
      poly(
        [
          [PLOT.left, PLOT.bottom],
          [PLOT.left, 14.6],
          [9.4, 8.6],
          [14, 12.4],
          [20.4, 4.6],
          [20.4, PLOT.bottom],
        ],
        true,
      ),
      PALETTE.accentSoft,
    ),
    stroked('M3.6 14.6 9.4 8.6 14 12.4 20.4 4.6', PALETTE.accent, 'bold'),
  ],
  chartPie: [
    filled(
      sector(PIE_CENTER, PIE_CENTER, PIE_RADIUS, Math.PI / 4, (3 * Math.PI) / 2),
      PALETTE.accent,
    ),
    filled(sector(PIE_CENTER, PIE_CENTER, PIE_RADIUS, -Math.PI / 2, 0), PALETTE.info),
    filled(sector(PIE_CENTER, PIE_CENTER, PIE_RADIUS, 0, Math.PI / 4), PALETTE.alt),
    stroked(
      join(
        line([PIE_CENTER, PIE_CENTER], piePoint(-Math.PI / 2)),
        line([PIE_CENTER, PIE_CENTER], piePoint(0)),
        line([PIE_CENTER, PIE_CENTER], piePoint(Math.PI / 4)),
      ),
      PALETTE.paper,
      'thin',
      { cap: 'butt' },
    ),
    stroked(circle(PIE_CENTER, PIE_CENTER, PIE_RADIUS), PALETTE.accentDeep, 'thin'),
  ],
  chartScatter: [
    axes(),
    filled(join(circle(7.4, 15.6, 1.7), circle(11.4, 10.4, 1.7)), PALETTE.info),
    filled(join(circle(15.2, 13.6, 1.7), circle(18.8, 6.6, 1.7)), PALETTE.alt),
  ],
} satisfies Record<string, IconDefinition>;
