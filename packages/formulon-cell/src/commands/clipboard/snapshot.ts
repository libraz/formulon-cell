import type { CellValue, Range } from '../../engine/types.js';
import { addrKey } from '../../engine/workbook-handle.js';
import type { CellFormat, State } from '../../store/store.js';

/**
 * Structured snapshot of a range — values, formulas, AND formats. Captured
 * on copy/cut so a subsequent Paste Special can pick which of the three
 * layers to apply. Spreadsheets keep this internal-clipboard separate from the
 * system clipboard; we mirror that.
 */
export interface ClipboardCell {
  formula: string | null;
  value: CellValue;
  format: CellFormat | undefined;
}

export interface ClipboardSnapshot {
  /** Original source range (sheet-relative). */
  range: Range;
  rows: number;
  cols: number;
  /** rows × cols matrix in row-major order. Empty source cells are present
   *  but with `value = { kind: 'blank' }` and undefined format. */
  cells: ClipboardCell[][];
}

const blankCell = (): ClipboardCell => ({
  formula: null,
  value: { kind: 'blank' },
  format: undefined,
});

export function captureSnapshot(state: State, range: Range): ClipboardSnapshot | null {
  const rows = range.r1 - range.r0 + 1;
  const cols = range.c1 - range.c0 + 1;
  if (rows <= 0 || cols <= 0) return null;
  // Cap at ~1M cells like the TSV copy path.
  if (rows * cols > 1_000_000) return null;

  const sheet = range.sheet;
  const grid: ClipboardCell[][] = [];
  for (let r = 0; r < rows; r += 1) {
    const line: ClipboardCell[] = [];
    for (let c = 0; c < cols; c += 1) {
      const key = addrKey({ sheet, row: range.r0 + r, col: range.c0 + c });
      const cell = state.data.cells.get(key);
      const fmt = state.format.formats.get(key);
      if (!cell && !fmt) {
        line.push(blankCell());
        continue;
      }
      line.push({
        formula: cell?.formula ?? null,
        value: cell?.value ?? { kind: 'blank' },
        format: fmt ? { ...fmt, borders: fmt.borders ? { ...fmt.borders } : undefined } : undefined,
      });
    }
    grid.push(line);
  }
  return { range, rows, cols, cells: grid };
}
