/** PivotTable field-pane area icons: the column and row axes as bar triplets. */

import { join, rect } from '../path.js';
import { filled } from '../primitives.js';
import { CENTER, OPTICAL, PALETTE } from '../tokens.js';
import type { IconDefinition } from '../types.js';

const BAR = 3;
const BAR_GAP = 4;
const BAR_LENGTH = OPTICAL * 0.9;

/** Three solid bars across the centre, standing (columns) or lying (rows). */
const barTriplet = (standing: boolean): IconDefinition => {
  const span = BAR * 3 + BAR_GAP * 2;
  const start = CENTER - span / 2;
  const long = CENTER - BAR_LENGTH / 2;
  const bars = [0, 1, 2].map((i) => {
    const offset = start + i * (BAR + BAR_GAP);
    return standing ? rect(offset, long, BAR, BAR_LENGTH) : rect(long, offset, BAR_LENGTH, BAR);
  });
  return [filled(join(...bars), PALETTE.ink)];
};

export const PIVOT_ICONS = {
  pivotAreaColumns: barTriplet(true),
  pivotAreaRows: barTriplet(false),
} satisfies Record<string, IconDefinition>;
