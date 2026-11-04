import type { State } from '../store/types.js';

type FilterVisibleState = Pick<State, 'ui' | 'layout'>;

/** Whether an AutoFilter on `sheet` hides `row`: a hidden data row inside a filter range. */
export const rowHiddenByFilter = (state: FilterVisibleState, sheet: number, row: number): boolean =>
  state.layout.hiddenRows.has(row) &&
  state.ui.filterRanges.some((r) => r.sheet === sheet && row > r.r0 && row <= r.r1);

/** The rows of `r0..r1` that no AutoFilter hides, in order. */
export const rowsVisibleToFilter = (
  state: FilterVisibleState,
  sheet: number,
  r0: number,
  r1: number,
): number[] => {
  const rows: number[] = [];
  for (let row = r0; row <= r1; row += 1) {
    if (!rowHiddenByFilter(state, sheet, row)) rows.push(row);
  }
  return rows;
};

/** `r0..r1` split into runs of rows that no AutoFilter hides. */
export const filterVisibleRowRuns = (
  state: FilterVisibleState,
  sheet: number,
  r0: number,
  r1: number,
): [number, number][] => {
  if (state.layout.hiddenRows.size === 0) return r1 >= r0 ? [[r0, r1]] : [];
  const runs: [number, number][] = [];
  for (const row of rowsVisibleToFilter(state, sheet, r0, r1)) {
    const last = runs[runs.length - 1];
    if (last && last[1] === row - 1) last[1] = row;
    else runs.push([row, row]);
  }
  return runs;
};
