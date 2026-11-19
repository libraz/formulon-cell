import { syncHiddenRowsToEngine } from '../engine/auto-filter-sync.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { sameRange, sameRangeList } from '../store/selection-geometry.js';
import type {
  FilterSortSpec,
  SpreadsheetStore,
  State,
  ValueFilterCriteria,
} from '../store/store.js';
import type { History } from './history.js';
import { sameJsonValue } from './prepared-change.js';

interface FilterSnapshot {
  hiddenRows: Set<number>;
  filterRanges: Range[];
  filterCriteria: ValueFilterCriteria[];
  filterSorts: FilterSortSpec[];
}

const cloneSorts = (sorts: readonly FilterSortSpec[] | undefined): FilterSortSpec[] =>
  (sorts ?? []).map((spec) => ({
    range: { ...spec.range },
    keys: spec.keys.map((key) => ({ ...key })),
  }));

const cloneRanges = (ranges: readonly Range[]): Range[] => ranges.map((range) => ({ ...range }));
const cloneCriteria = (criteria: readonly ValueFilterCriteria[]): ValueFilterCriteria[] =>
  criteria.map((c) => ({
    range: { ...c.range },
    byCol: c.byCol,
    hiddenValues: [...c.hiddenValues],
    ...(c.condition ? { condition: { ...c.condition } } : {}),
    ...(c.color ? { color: { ...c.color } } : {}),
  }));

function captureFilterSnapshot(state: State): FilterSnapshot {
  return {
    hiddenRows: new Set(state.layout.hiddenRows),
    filterRanges: cloneRanges(state.ui.filterRanges),
    filterCriteria: cloneCriteria(state.ui.filterCriteria),
    filterSorts: cloneSorts(state.ui.filterSorts),
  };
}

function applyFilterSnapshot(
  store: SpreadsheetStore,
  snap: FilterSnapshot,
  wb?: WorkbookHandle,
): void {
  const beforeHidden = new Set(store.getState().layout.hiddenRows);
  store.setState((s) => ({
    ...s,
    layout: { ...s.layout, hiddenRows: new Set(snap.hiddenRows) },
    ui: {
      ...s.ui,
      filterRanges: cloneRanges(snap.filterRanges),
      filterCriteria: cloneCriteria(snap.filterCriteria),
      filterSorts: cloneSorts(snap.filterSorts),
    },
  }));
  if (wb) {
    syncHiddenRowsToEngine(wb, store.getState().data.sheetIndex, beforeHidden, snap.hiddenRows);
  }
}

function sameFilterSnapshot(a: FilterSnapshot, b: FilterSnapshot): boolean {
  if (!sameRangeList(a.filterRanges, b.filterRanges)) return false;
  if (a.hiddenRows.size !== b.hiddenRows.size) return false;
  if (!sameCriteriaList(a.filterCriteria, b.filterCriteria)) return false;
  if (!sameJsonValue(a.filterSorts, b.filterSorts)) return false;
  for (const row of a.hiddenRows) {
    if (!b.hiddenRows.has(row)) return false;
  }
  return true;
}

function sameCriteriaList(
  a: readonly ValueFilterCriteria[],
  b: readonly ValueFilterCriteria[],
): boolean {
  if (a.length !== b.length) return false;
  return a.every((left, i) => {
    const right = b[i];
    return (
      !!right &&
      left.byCol === right.byCol &&
      sameRange(left.range, right.range) &&
      sameJsonValue(left.condition ?? null, right.condition ?? null) &&
      sameJsonValue(left.color ?? null, right.color ?? null) &&
      left.hiddenValues.length === right.hiddenValues.length &&
      left.hiddenValues.every((value, index) => value === right.hiddenValues[index])
    );
  });
}

export function recordFilterChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  mutate: () => T,
  wb?: WorkbookHandle,
): T {
  const before = captureFilterSnapshot(store.getState());
  const result = mutate();
  const after = captureFilterSnapshot(store.getState());
  if (history && !history.isReplaying() && !sameFilterSnapshot(before, after)) {
    history.push({
      undo: () => applyFilterSnapshot(store, before, wb),
      redo: () => applyFilterSnapshot(store, after, wb),
    });
  }
  return result;
}
