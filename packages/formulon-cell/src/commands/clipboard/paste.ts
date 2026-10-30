import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { State } from '../../store/store.js';
import {
  coercedInputToPatch,
  coerceInputForCell,
  inputLocaleForWorkbook,
  writeCoerced,
} from '../coerce-input.js';
import { blockDataTableEdit, blockTableFunction } from '../data-table-guard.js';
import type { History } from '../history.js';
import { blockPivotEdit } from '../pivot-guard.js';
import { isCellWritable } from '../protection.js';
import { applyPatchesAtomic } from './apply-patches-atomic.js';
import { parseTSV } from './tsv.js';

export interface PasteResult {
  writtenRange: Range;
}

/**
 * Paste a TSV payload at the current active cell. Each value is run through
 * `coerceInput` so leading-`=` strings become formulas, numerics become
 * numbers, etc. — matching the spreadsheet behaviour when you paste values from a
 * different program. Cells gated by sheet protection (locked + sheet
 * protected) are silently skipped, mirroring the spreadsheet's paste behaviour.
 * With `history`, the cells are written in one atomic engine call and recorded
 * as a single undo step; without it, each cell goes through the scalar path.
 */
export function pasteTSV(
  state: State,
  wb: WorkbookHandle,
  text: string,
  history: History | null = null,
): PasteResult | null {
  if (!text) return null;
  const rows = parseTSV(text);
  if (rows.length === 0) return null;

  const origin: Addr = state.selection.active;
  const sheet = origin.sheet;
  const pastedCols = rows.reduce((widest, row) => Math.max(widest, row.length), 1);
  const pasteArea: Range = {
    sheet,
    r0: origin.row,
    c0: origin.col,
    r1: origin.row + rows.length - 1,
    c1: origin.col + pastedCols - 1,
  };
  if (blockPivotEdit(wb, pasteArea) || blockDataTableEdit(wb, pasteArea)) return null;
  const inputLocale = inputLocaleForWorkbook(wb);
  let maxCols = 0;
  const writes: { addr: Addr; input: ReturnType<typeof coerceInputForCell> }[] = [];
  for (let r = 0; r < rows.length; r += 1) {
    const cells = rows[r] ?? [];
    if (cells.length > maxCols) maxCols = cells.length;
    for (let c = 0; c < cells.length; c += 1) {
      const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
      if (!isCellWritable(state, addr)) continue;
      writes.push({
        addr,
        input: coerceInputForCell(state, addr, cells[c] ?? '', inputLocale, wb),
      });
    }
  }
  const formulas = writes.map(({ input }) => (input.kind === 'formula' ? input.text : null));
  if (blockTableFunction(wb, formulas)) return null;

  if (history) {
    applyPatchesAtomic(
      wb,
      writes.map(({ addr, input }) => coercedInputToPatch(addr, input)),
      history,
    );
  } else {
    wb.withBatchedRecalc(() => {
      for (const { addr, input } of writes) writeCoerced(wb, addr, input);
    });
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
