import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { SpreadsheetStore, State } from '../store/store.js';

export type FilterPredicate = (
  cell: { value: unknown; formula: string | null } | undefined,
) => boolean;

/** Hide rows in `range` whose `byCol` value fails `predicate`. The first
 *  row of the range is treated as a header and stays visible.
 *  Hidden rows are added to `layout.hiddenRows` (UI-only — engine state is
 *  untouched). The range is also stamped onto `ui.filterRange` so the column
 *  headers paint the autofilter chevron. */
export function applyFilter(
  state: State,
  store: SpreadsheetStore,
  range: Range,
  byCol: number,
  predicate: FilterPredicate,
): number {
  const newHidden = new Set<number>(state.layout.hiddenRows);
  let hiddenCount = 0;
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: byCol }));
    if (!predicate(cell)) {
      if (!newHidden.has(r)) {
        newHidden.add(r);
        hiddenCount += 1;
      }
    }
  }
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: newHidden },
    ui: { ...s.ui, filterRange: { ...range } },
  }));
  return hiddenCount;
}

/** Mark `range` as the active autofilter region without applying any predicate.
 *  Headers in the range paint the chevron so the user can open the dropdown
 *  per column. Equivalent to the "Filter" toggle (Ctrl+Shift+L). */
export function setAutoFilter(store: SpreadsheetStore, range: Range | null): void {
  store.setState((s) => ({ ...s, ui: { ...s.ui, filterRange: range ? { ...range } : null } }));
}

/** Reveal all rows in `range` (or all rows when range is omitted). Also clears
 *  `ui.filterRange` when `range` matches the active autofilter region (or when
 *  no range is supplied). */
export function clearFilter(state: State, store: SpreadsheetStore, range?: Range): void {
  const next = new Set<number>(state.layout.hiddenRows);
  if (!range) {
    next.clear();
  } else {
    for (let r = range.r0; r <= range.r1; r += 1) next.delete(r);
  }
  store.setState((s) => {
    const fr = s.ui.filterRange;
    const clearFr =
      !range ||
      (fr != null &&
        fr.sheet === range.sheet &&
        fr.r0 === range.r0 &&
        fr.r1 === range.r1 &&
        fr.c0 === range.c0 &&
        fr.c1 === range.c1);
    return {
      ...s,
      layout: { ...s.layout, hiddenRows: next },
      ui: clearFr ? { ...s.ui, filterRange: null } : s.ui,
    };
  });
}

/** Distinct (string-coerced) values found in `byCol` of `range`, excluding
 *  the header row. Useful for building a dropdown of filter options. */
export function distinctValues(state: State, range: Range, byCol: number): string[] {
  const out = new Set<string>();
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const cell = state.data.cells.get(addrKey({ sheet: range.sheet, row: r, col: byCol }));
    if (!cell) {
      out.add('');
      continue;
    }
    const v = cell.value;
    if (v.kind === 'number') out.add(String(v.value));
    else if (v.kind === 'text') out.add(v.value);
    else if (v.kind === 'bool') out.add(v.value ? 'TRUE' : 'FALSE');
    else out.add('');
  }
  return Array.from(out).sort();
}
