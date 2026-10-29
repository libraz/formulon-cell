import { mergeAnchorOf, mergeAt } from '../commands/merge.js';
import type { Addr, Range } from '../engine/types.js';
import { isRowHidden } from '../store/row-visibility.js';
import type { LayoutSlice, State } from '../store/types.js';

export type VisibleAxis = 'row' | 'col';

type IndexSet = ReadonlySet<number>;

const sortedCache = new WeakMap<object, number[]>();
const EMPTY_SET: ReadonlySet<number> = new Set<number>();

function sorted(set: IndexSet): number[] {
  let result = sortedCache.get(set as object);
  if (!result) {
    result = [...set].sort((a, b) => a - b);
    sortedCache.set(set as object, result);
  }
  return result;
}

function lowerBound(values: readonly number[], value: number): number {
  let lo = 0;
  let hi = values.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    const item = values[mid];
    if (item === undefined || item < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function axisSet(layout: LayoutSlice, axis: VisibleAxis): IndexSet {
  return axis === 'row' ? layout.hiddenRows : layout.hiddenCols;
}

function indexIsHidden(layout: LayoutSlice, axis: VisibleAxis, index: number): boolean {
  return axis === 'row' ? isRowHidden(layout, index) : layout.hiddenCols.has(index);
}

export const isVisibleAxisIndex = (
  layout: LayoutSlice,
  axis: VisibleAxis,
  index: number,
): boolean => !indexIsHidden(layout, axis, index);

export function isCellVisible(layout: LayoutSlice, row: number, col: number): boolean {
  return !indexIsHidden(layout, 'row', row) && !indexIsHidden(layout, 'col', col);
}

/** Return the closest visible index strictly after/before `from`, within bounds. */
export function nextVisibleIndex(
  layout: LayoutSlice,
  axis: VisibleAxis,
  from: number,
  direction: -1 | 1,
  min: number,
  max: number,
): number | null {
  const first = direction > 0 ? Math.max(min, from + 1) : Math.min(max, from - 1);
  if (first < min || first > max) return null;

  if (axis === 'row' && layout.defaultRowsHidden === true) {
    const shown = sorted(layout.shownRows ?? EMPTY_SET);
    let i = direction > 0 ? lowerBound(shown, first) : lowerBound(shown, first + 1) - 1;
    while (i >= 0 && i < shown.length) {
      const row = shown[i];
      if (row === undefined) break;
      if ((direction > 0 && row > max) || (direction < 0 && row < min)) return null;
      if (!layout.hiddenRows.has(row)) return row;
      i += direction;
    }
    return null;
  }

  const hidden = sorted(axisSet(layout, axis));
  let cursor = first;
  if (direction > 0) {
    for (let i = lowerBound(hidden, first); i < hidden.length; i += 1) {
      const index = hidden[i];
      if (index === undefined || index > max) break;
      if (index > cursor) return cursor;
      if (index === cursor) cursor += 1;
    }
    return cursor <= max ? cursor : null;
  }
  for (let i = lowerBound(hidden, first + 1) - 1; i >= 0; i -= 1) {
    const index = hidden[i];
    if (index === undefined || index < min) break;
    if (index < cursor) return cursor;
    if (index === cursor) cursor -= 1;
  }
  return cursor >= min ? cursor : null;
}

/** First visible index in a bounded axis interval, in the requested order. */
export function firstVisibleIndex(
  layout: LayoutSlice,
  axis: VisibleAxis,
  min: number,
  max: number,
  direction: -1 | 1,
): number | null {
  return nextVisibleIndex(layout, axis, direction > 0 ? min - 1 : max + 1, direction, min, max);
}

/** Count visible indexes without walking every row or column in a large range. */
export function visibleAxisCount(
  layout: LayoutSlice,
  axis: VisibleAxis,
  min: number,
  max: number,
): number {
  if (min > max) return 0;
  if (axis === 'row' && layout.defaultRowsHidden === true) {
    const shown = sorted(layout.shownRows ?? EMPTY_SET);
    const shownStart = lowerBound(shown, min);
    const shownEnd = lowerBound(shown, max + 1);
    const shownCount = shownEnd - shownStart;
    const hidden = sorted(layout.hiddenRows);
    const hiddenStart = lowerBound(hidden, min);
    const hiddenEnd = lowerBound(hidden, max + 1);
    const hiddenCount = hiddenEnd - hiddenStart;
    if (shownCount <= hiddenCount) {
      let count = 0;
      for (let i = shownStart; i < shownEnd; i += 1) {
        const row = shown[i];
        if (row !== undefined && !layout.hiddenRows.has(row)) count += 1;
      }
      return count;
    }
    let excluded = 0;
    for (let i = hiddenStart; i < hiddenEnd; i += 1) {
      const row = hidden[i];
      if (row !== undefined && layout.shownRows?.has(row)) excluded += 1;
    }
    return shownCount - excluded;
  }
  const hidden = sorted(axisSet(layout, axis));
  const hiddenCount = lowerBound(hidden, max + 1) - lowerBound(hidden, min);
  return Math.max(0, max - min + 1 - hiddenCount);
}

/** Count visible cells after collapsing each intersecting merge to one stop. */
export function visibleLogicalStopCount(
  layout: LayoutSlice,
  range: Range,
  merges: Iterable<Range>,
): number {
  const rows = visibleAxisCount(layout, 'row', range.r0, range.r1);
  const cols = visibleAxisCount(layout, 'col', range.c0, range.c1);
  let stops = rows * cols;
  for (const merge of merges) {
    if (merge.sheet !== range.sheet) continue;
    const r0 = Math.max(range.r0, merge.r0);
    const r1 = Math.min(range.r1, merge.r1);
    const c0 = Math.max(range.c0, merge.c0);
    const c1 = Math.min(range.c1, merge.c1);
    if (r0 > r1 || c0 > c1) continue;
    const mergedVisibleCells =
      visibleAxisCount(layout, 'row', r0, r1) * visibleAxisCount(layout, 'col', c0, c1);
    if (mergedVisibleCells > 0) stops -= mergedVisibleCells - 1;
  }
  return stops;
}

/**
 * Move by one visible row/column. A visible step into a merge selects its
 * anchor; a hidden step inside one skips beyond that merge. `stepWithMerge`
 * stays the geometric step for callers that ignore visibility.
 */
export function stepVisibleWithMerge(
  state: State,
  from: Addr,
  dRow: number,
  dCol: number,
  maxRow: number,
  maxCol: number,
  minRow = 0,
  minCol = 0,
): Addr {
  const axis: VisibleAxis = dRow === 0 ? 'col' : 'row';
  const direction = (dRow || dCol) > 0 ? 1 : -1;
  const min = axis === 'row' ? minRow : minCol;
  const max = axis === 'row' ? maxRow : maxCol;
  const indexOf = (addr: Addr): number => (axis === 'row' ? addr.row : addr.col);
  const atIndex = (index: number): Addr =>
    axis === 'row' ? { ...from, row: index } : { ...from, col: index };
  const beyondMerge = (merge: Range): number =>
    axis === 'row'
      ? direction > 0
        ? merge.r1 + 1
        : merge.r0 - 1
      : direction > 0
        ? merge.c1 + 1
        : merge.c0 - 1;

  const sourceMerge = mergeAt(state, from);
  let destinationIndex: number;
  if (sourceMerge) {
    destinationIndex = beyondMerge(sourceMerge);
  } else {
    const rawIndex = Math.max(min, Math.min(max, indexOf(from) + direction));
    const rawCandidate = atIndex(rawIndex);
    const rawMerge = mergeAt(state, rawCandidate);
    if (rawMerge && indexIsHidden(state.layout, axis, rawIndex)) {
      destinationIndex = beyondMerge(rawMerge);
    } else if (indexIsHidden(state.layout, axis, rawIndex)) {
      const next = nextVisibleIndex(state.layout, axis, rawIndex, direction, min, max);
      if (next === null) return from;
      const landed = atIndex(next);
      const landedMerge = mergeAt(state, landed);
      const anchorIndex = axis === 'row' ? landedMerge?.r0 : landedMerge?.c0;
      const hiddenLeadingAnchor =
        direction > 0 &&
        landedMerge !== null &&
        anchorIndex !== undefined &&
        indexIsHidden(state.layout, axis, anchorIndex);
      destinationIndex = hiddenLeadingAnchor ? beyondMerge(landedMerge) : next;
    } else {
      destinationIndex = rawIndex;
    }
  }

  if (destinationIndex < min || destinationIndex > max) return from;
  if (indexIsHidden(state.layout, axis, destinationIndex)) {
    const visible = nextVisibleIndex(
      state.layout,
      axis,
      destinationIndex - direction,
      direction,
      min,
      max,
    );
    if (visible === null) return from;
    destinationIndex = visible;
  }

  const destination = atIndex(destinationIndex);
  return mergeAt(state, destination) ? mergeAnchorOf(state, destination) : destination;
}
