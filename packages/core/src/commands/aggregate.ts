import type { State } from '../store/store.js';

export interface SelectionStats {
  /** Total cells in selection, including blanks. */
  cells: number;
  /** Cells whose value.kind === 'number'. */
  numericCount: number;
  /** Cells with any non-blank value (Excel "Count"). */
  nonBlankCount: number;
  sum: number;
  /** Only meaningful when numericCount > 0. */
  avg: number;
  min: number;
  max: number;
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
 * not twice — Excel parity.
 */
export function aggregateSelection(state: State): SelectionStats {
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const sheet = state.data.sheetIndex;

  let cells = 0;
  for (const r of ranges) {
    if (r.r1 < r.r0 || r.c1 < r.c0) continue;
    cells += (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);
  }
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
