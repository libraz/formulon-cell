import type { Range } from '../engine/types.js';
import { isRowHidden } from '../store/row-visibility.js';
import { sameRange } from '../store/selection-geometry.js';
import type { State } from '../store/types.js';

/** The AutoFilter ranges of the current sheet that have at least one column criterion. */
export function activeFilterRanges(
  state: Pick<State, 'ui' | 'data'>,
): Array<{ range: Range; cols: Set<number> }> {
  const sheet = state.data.sheetIndex;
  const out: Array<{ range: Range; cols: Set<number> }> = [];
  for (const range of state.ui.filterRanges) {
    if (range.sheet !== sheet) continue;
    const cols = new Set<number>();
    for (const c of state.ui.filterCriteria) {
      if (sameRange(c.range, range)) cols.add(c.byCol);
    }
    if (cols.size > 0) out.push({ range, cols });
  }
  return out;
}

/** Whether a column of the sheet's AutoFilter carries a criterion. */
export function filterColumnActive(
  state: Pick<State, 'ui' | 'data'>,
  range: Range,
  col: number,
): boolean {
  return activeFilterRanges(state).some((f) => sameRange(f.range, range) && f.cols.has(col));
}

/** Data rows of every filtered range that survive the filter, plus how many exist. */
export function filterRecordCounts(
  state: Pick<State, 'ui' | 'data' | 'layout'>,
): { found: number; total: number } | null {
  const active = activeFilterRanges(state);
  if (active.length === 0) return null;
  let found = 0;
  let total = 0;
  for (const { range } of active) {
    for (let r = range.r0 + 1; r <= range.r1; r += 1) {
      total += 1;
      if (!isRowHidden(state.layout, r)) found += 1;
    }
  }
  return { found, total };
}
