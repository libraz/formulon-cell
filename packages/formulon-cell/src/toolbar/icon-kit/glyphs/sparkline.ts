/** Sparkline Design icons: a trace or a bar strip, with the points a command picks out. */

import { circle, hLine, join, rect } from '../path.js';
import { filled, stroked } from '../primitives.js';
import { PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

const TRACE = [
  [2.8, 12],
  [7.4, 8],
  [11.6, 17],
  [16, 6],
  [21.2, 10],
] as const;

const trace = () =>
  stroked(`M${TRACE.map(([x, y]) => `${x} ${y}`).join(' ')}`, PALETTE.info, 'bold');

/** The trace with the points at `indexes` marked. */
const marked = (indexes: readonly number[], fill: string): IconDefinition => [
  trace(),
  filled(join(...indexes.map((i) => circle(TRACE[i]?.[0] ?? 0, TRACE[i]?.[1] ?? 0, 2))), fill),
];

const BAR_X = [3.4, 7.6, 11.8, 16, 20.2] as const;
const BAR_W = 3;
const MID = 12;

/** A strip of bars about the middle line; a negative bar hangs below it in red. */
const strip = (heights: readonly number[]): IconDefinition => [
  stroked(hLine(2.4, MID, 19.2), PALETTE.grid, 'hairline'),
  ...heights.map((h, i) =>
    filled(
      h >= 0 ? rect(BAR_X[i] ?? 0, MID - h, BAR_W, h) : rect(BAR_X[i] ?? 0, MID, BAR_W, -h),
      h >= 0 ? PALETTE.info : PALETTE.danger,
    ),
  ),
];

export const SPARKLINE_ICONS = {
  sparklineColumn: [
    ...[6, 9, 4, 11, 7].map((h, i) => filled(rect(BAR_X[i] ?? 0, 19 - h, BAR_W, h), PALETTE.info)),
  ],
  sparklineWinLoss: strip([6, -6, 6, 6, -6]),
  sparklineNegative: strip([6, 9, -5, 11, -7]),
  sparklineHigh: marked([3], PALETTE.danger),
  sparklineLow: marked([2], PALETTE.danger),
  sparklineFirst: marked([0], PALETTE.danger),
  sparklineLast: marked([4], PALETTE.danger),
  sparklineMarkers: marked([0, 1, 2, 3, 4], PALETTE.info),
  sparklineAxis: [trace(), stroked(hLine(2.4, 20, 19.2), PALETTE.ink, 'regular')],
} satisfies Record<string, IconDefinition>;
