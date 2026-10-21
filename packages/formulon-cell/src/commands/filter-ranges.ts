import type { Addr, Range } from '../engine/types.js';
import { rangeContainsAddr, rangesIntersect, sameRange } from '../store/selection-geometry.js';
import type { State } from '../store/store.js';

/**
 * Lookups over `ui.filterRanges`: a sheet can carry its own AutoFilter plus one
 * per table, all showing buttons at once. A range on a table overlay's header
 * row and columns is that table's filter; any other range is the worksheet's.
 * Data > Filter, Clear and Reapply act on the filter holding the active cell,
 * otherwise on the worksheet filter, as the desktop spreadsheet does.
 */

type FilterRangeState = Pick<State, 'ui' | 'tables'>;

/** True when `range` is a table's filter: it starts on the table's header row
 *  across its columns and ends inside it (a table filter skips a totals row). */
export function isTableFilterRange(state: Pick<State, 'tables'>, range: Range): boolean {
  return state.tables.tables.some(
    ({ range: t }) =>
      t.sheet === range.sheet &&
      t.r0 === range.r0 &&
      t.c0 === range.c0 &&
      t.c1 === range.c1 &&
      range.r1 <= t.r1,
  );
}

/** The filter range containing `addr`, or null. */
export function filterRangeAt(state: Pick<State, 'ui'>, addr: Addr): Range | null {
  return state.ui.filterRanges.find((range) => rangeContainsAddr(range, addr)) ?? null;
}

/** The filter range the Data-tab filter commands act on: the one holding the
 *  active cell, else the worksheet (non-table) filter on the active cell's sheet. */
export function activeFilterRange(
  state: FilterRangeState & Pick<State, 'selection'>,
): Range | null {
  const active = state.selection.active;
  return (
    filterRangeAt(state, active) ??
    state.ui.filterRanges.find(
      (range) => range.sheet === active.sheet && !isTableFilterRange(state, range),
    ) ??
    null
  );
}

/** Filter ranges on the active sheet, the active one first so a column shared
 *  by stacked filters opens the dropdown of the filter being worked on. */
export function headerFilterRanges(
  state: FilterRangeState & Pick<State, 'selection' | 'data'>,
): Range[] {
  const sheet = state.data.sheetIndex;
  const active = activeFilterRange(state);
  const ranges = state.ui.filterRanges.filter((range) => range.sheet === sheet);
  if (!active || active.sheet !== sheet) return ranges;
  return [active, ...ranges.filter((range) => range !== active)];
}

/** `filterRanges` with `range` added. It replaces every range it overlaps (the
 *  same filter grown or moved), and a worksheet filter also replaces the
 *  sheet's previous worksheet filter, since a sheet holds only one. */
export function withFilterRange(state: FilterRangeState, range: Range): Range[] {
  const worksheet = !isTableFilterRange(state, range);
  const kept = state.ui.filterRanges.filter(
    (existing) =>
      !rangesIntersect(existing, range) &&
      !(worksheet && existing.sheet === range.sheet && !isTableFilterRange(state, existing)),
  );
  return [...kept, { ...range }];
}

/** `filterRanges` without the entry exactly matching `range`. */
export function withoutFilterRange(ranges: readonly Range[], range: Range): Range[] {
  return ranges.filter((existing) => !sameRange(existing, range));
}
