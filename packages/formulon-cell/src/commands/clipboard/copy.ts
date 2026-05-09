import type { Range } from '../../engine/types.js';
import type { State } from '../../store/store.js';
import { encodeTSV } from './tsv.js';

export interface CopyResult {
  /** TSV payload — values only; formulas resolve to their displayed text. */
  tsv: string;
  /** The range that was copied. */
  range: Range;
}

/**
 * Snapshot the current selection into a TSV payload. Values come from the
 * store's cached cell map (no engine reads) — for formula cells this means
 * the last computed value, matching the spreadsheet's "values only" copy semantic.
 */
export function copy(state: State): CopyResult | null {
  const r = state.selection.range;
  const sheet = state.data.sheetIndex;
  if (r.r1 < r.r0 || r.c1 < r.c0) return null;
  // Refuse to materialise an entire-sheet copy — 17B blank strings would OOM.
  if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > 1_000_000) return null;

  const grid: string[][] = [];
  for (let row = r.r0; row <= r.r1; row += 1) {
    const line: string[] = [];
    for (let col = r.c0; col <= r.c1; col += 1) {
      line.push(displayValue(state, sheet, row, col));
    }
    grid.push(line);
  }
  return { tsv: encodeTSV(grid), range: r };
}

function displayValue(state: State, sheet: number, row: number, col: number): string {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return '';
  switch (cell.value.kind) {
    case 'number':
      return String(cell.value.value);
    case 'text':
      return cell.value.value;
    case 'bool':
      return cell.value.value ? 'TRUE' : 'FALSE';
    case 'error':
      return cell.value.text;
    default:
      return '';
  }
}
