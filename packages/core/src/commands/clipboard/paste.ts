import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { State } from '../../store/store.js';
import { coerceInput, writeCoerced } from '../coerce-input.js';
import { parseTSV } from './tsv.js';

export interface PasteResult {
  writtenRange: Range;
}

/**
 * Paste a TSV payload at the current active cell. Each value is run through
 * `coerceInput` so leading-`=` strings become formulas, numerics become
 * numbers, etc. — matching Excel's behaviour when you paste values from a
 * different program.
 */
export function pasteTSV(state: State, wb: WorkbookHandle, text: string): PasteResult | null {
  if (!text) return null;
  const rows = parseTSV(text);
  if (rows.length === 0) return null;

  const origin: Addr = state.selection.active;
  const sheet = origin.sheet;
  let maxCols = 0;

  for (let r = 0; r < rows.length; r += 1) {
    const cells = rows[r] ?? [];
    if (cells.length > maxCols) maxCols = cells.length;
    for (let c = 0; c < cells.length; c += 1) {
      const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
      writeCoerced(wb, addr, coerceInput(cells[c] ?? ''));
    }
  }

  return {
    writtenRange: {
      sheet,
      r0: origin.row,
      c0: origin.col,
      r1: origin.row + rows.length - 1,
      c1: origin.col + Math.max(0, maxCols - 1),
    },
  };
}
