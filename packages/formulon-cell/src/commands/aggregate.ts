import type { Range } from '../engine/types.js';
import type { State, StatusAggKey } from '../store/store.js';

export const STATUS_AGGREGATE_KEYS: readonly StatusAggKey[] = [
  'average',
  'count',
  'countNumbers',
  'min',
  'max',
  'sum',
];

export interface SelectionStats {
  /** Total cells in selection, including blanks. */
  cells: number;
  /** Cells whose value.kind === 'number'. */
  numericCount: number;
  /** Cells with any non-blank value (the "Count" aggregate). */
  nonBlankCount: number;
  sum: number;
  /** Only meaningful when numericCount > 0. */
  avg: number;
  min: number;
  max: number;
}

export interface StatusAggregateEntry {
  key: StatusAggKey;
  value: number;
}

const EMPTY: SelectionStats = {
  cells: 0,
  numericCount: 0,
  nonBlankCount: 0,
  sum: 0,
  avg: 0,
  min: 0,
  max: 0,
};

/**
 * Compute the standard Sum / Avg / Count for the current selection from the
 * cached cell map. Pure & engine-free — never triggers a recalc. Status bar
 * subscribers can call this on every selection change without overhead.
 *
 * Multi-range selections (`selection.extraRanges`) are summed alongside the
 * primary range. Cells inside an overlap between two ranges are counted once,
 * not twice — spreadsheet parity.
 */
export function aggregateSelection(state: State): SelectionStats {
  const sheet = state.data.sheetIndex;
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])].filter(
    (range) => range.sheet === sheet,
  );

  const cells = countUniqueRangeCells(ranges);
  if (cells <= 0) return EMPTY;

  let numericCount = 0;
  let nonBlankCount = 0;
  let sum = 0;
  let min = Number.POSITIVE_INFINITY;
  let max = Number.NEGATIVE_INFINITY;

  const visited = new Set<string>();
  // Iterate the populated cell map rather than the range. selectAll() can hand
  // us a 17B-cell rectangle; the cell map is bounded by what the user actually
  // typed and stays cheap.
  for (const [key, cell] of state.data.cells) {
    const parts = key.split(':');
    if (parts.length !== 3) continue;
    if (Number(parts[0]) !== sheet) continue;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    let inAny = false;
    for (const r of ranges) {
      if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1) continue;
      inAny = true;
      break;
    }
    if (!inAny) continue;
    if (visited.has(key)) continue;
    visited.add(key);
    if (cell.value.kind === 'blank') continue;
    nonBlankCount += 1;
    if (cell.value.kind !== 'number') continue;
    const n = cell.value.value;
    numericCount += 1;
    sum += n;
    if (n < min) min = n;
    if (n > max) max = n;
  }

  if (numericCount === 0) {
    return { cells, numericCount: 0, nonBlankCount, sum: 0, avg: 0, min: 0, max: 0 };
  }
  return { cells, numericCount, nonBlankCount, sum, avg: sum / numericCount, min, max };
}

export function statusAggregateValue(key: StatusAggKey, stats: SelectionStats): number | null {
  if (key === 'count') return stats.nonBlankCount > 0 ? stats.nonBlankCount : null;
  if (key === 'countNumbers') return stats.numericCount > 0 ? stats.numericCount : null;
  if (stats.numericCount === 0) return null;
  switch (key) {
    case 'sum':
      return stats.sum;
    case 'average':
      return stats.avg;
    case 'min':
      return stats.min;
    case 'max':
      return stats.max;
    default:
      return null;
  }
}

export function visibleStatusAggregates(state: State): readonly StatusAggregateEntry[] {
  const stats = aggregateSelection(state);
  const out: StatusAggregateEntry[] = [];
  for (const key of state.ui.statusAggs) {
    const value = statusAggregateValue(key, stats);
    if (value != null) out.push({ key, value });
  }
  return out;
}

export function countUniqueRangeCells(ranges: readonly Range[]): number {
  const normalized = ranges
    .filter((r) => r.r1 >= r.r0 && r.c1 >= r.c0)
    .map((r) => ({ r0: r.r0, r1: r.r1, c0: r.c0, c1: r.c1 }));
  if (normalized.length === 0) return 0;

  const boundaries = new Set<number>();
  for (const r of normalized) {
    boundaries.add(r.r0);
    boundaries.add(r.r1 + 1);
  }
  const rows = [...boundaries].sort((a, b) => a - b);
  let total = 0;
  for (let i = 0; i < rows.length - 1; i += 1) {
    const start = rows[i];
    const end = rows[i + 1];
    if (start === undefined || end === undefined || end <= start) continue;
    const intervals: [number, number][] = [];
    for (const r of normalized) {
      if (r.r0 <= start && r.r1 + 1 >= end) intervals.push([r.c0, r.c1]);
    }
    total += (end - start) * countUniqueColumns(intervals);
  }
  return total;
}

const countUniqueColumns = (intervals: [number, number][]): number => {
  if (intervals.length === 0) return 0;
  intervals.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  let total = 0;
  let [start, end] = intervals[0] ?? [0, -1];
  for (let i = 1; i < intervals.length; i += 1) {
    const next = intervals[i];
    if (!next) continue;
    if (next[0] <= end + 1) {
      end = Math.max(end, next[1]);
      continue;
    }
    total += end - start + 1;
    [start, end] = next;
  }
  total += end - start + 1;
  return total;
};
