import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { CellFormat, SpreadsheetStore, State } from '../../store/store.js';
import {
  coercedInputToPatch,
  coerceInputForCell,
  inputLocaleForWorkbook,
  writeCoerced,
} from '../coerce-input.js';
import {
  blockDataTableEdit,
  blockTableFunction,
  dataTableGuardAlert,
  tableFunctionAlert,
} from '../data-table-guard.js';
import { type History, transactIn } from '../history.js';
import { appliedBatch, type ChangeBatchResult, refusedBatch } from '../interaction-policy.js';
import { blockPivotEdit, pivotGuardAlert } from '../pivot-guard.js';
import { isCellWritable, isSheetProtected } from '../protection.js';
import { recordFormatChange } from '../slice-history.js';
import { applyPatchesAtomic } from './apply-patches-atomic.js';
import { extendTablesForPaste } from './table-extension.js';
import { parseTSV } from './tsv.js';

/**
 * Paste a TSV payload at the current active cell. Each value is run through
 * `coerceInput` so leading-`=` strings become formulas, numerics become
 * numbers, etc. — matching the spreadsheet behaviour when you paste values from a
 * different program. A destination that touches a locked cell of a protected
 * sheet refuses the paste as a whole, as the spreadsheet does.
 * With `history`, the cells are written in one atomic engine call and recorded
 * as a single undo step; without it, each cell goes through the scalar path.
 *
 * Like Excel's, the paste also shapes the sheet around the values: a General
 * cell takes the number format its text implies (`1,234` #,##0, a date,
 * `12:30` h:mm, `1e3` scientific), and a table the paste lands against grows
 * over it (see {@link extendTablesForPaste}).
 */
export function pasteTSV(
  state: State,
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  text: string,
  history: History | null = null,
): ChangeBatchResult {
  const rows = text ? parseTSV(text) : [];
  if (rows.length === 0) return refusedBatch({ code: 'invalid', reason: 'nothing to paste' });

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
  if (!rangeWritable(state, pasteArea)) return refusedBatch({ code: 'protected' });
  if (blockPivotEdit(wb, pasteArea, 'paste'))
    return refusedBatch({ code: 'protected', reason: pivotGuardAlert(wb, 'paste').message });
  if (blockDataTableEdit(wb, pasteArea))
    return refusedBatch({ code: 'protected', reason: dataTableGuardAlert(wb).message });
  const inputLocale = inputLocaleForWorkbook(wb);
  let maxCols = 0;
  const writes: { addr: Addr; input: ReturnType<typeof coerceInputForCell> }[] = [];
  for (let r = 0; r < rows.length; r += 1) {
    const cells = rows[r] ?? [];
    if (cells.length > maxCols) maxCols = cells.length;
    for (let c = 0; c < cells.length; c += 1) {
      const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
      writes.push({
        addr,
        input: coerceInputForCell(state, addr, cells[c] ?? '', inputLocale, wb),
      });
    }
  }
  const formulas = writes.map(({ input }) => (input.kind === 'formula' ? input.text : null));
  if (blockTableFunction(wb, formulas))
    return refusedBatch({ code: 'invalid', reason: tableFunctionAlert(wb).message });

  return transactIn(history, () => {
    const { changed } = wb.trackChangedCells(() => {
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
    });

    const writtenRange: Range = {
      sheet,
      r0: origin.row,
      c0: origin.col,
      r1: origin.row + rows.length - 1,
      c1: origin.col + Math.max(0, maxCols - 1),
    };
    applyImplicitFormats(store, history, writes);
    extendTablesForPaste(store, history, writtenRange);
    return appliedBatch(changed, writtenRange);
  });
}

/** Whether every cell of `range` is writable under sheet protection. */
function rangeWritable(state: State, range: Range): boolean {
  if (!isSheetProtected(state, range.sheet)) return true;
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (!isCellWritable(state, { sheet: range.sheet, row, col })) return false;
    }
  }
  return true;
}

/** Give each General destination the number format its pasted text implies;
 *  an explicit format stays. */
function applyImplicitFormats(
  store: SpreadsheetStore,
  history: History | null,
  writes: readonly { addr: Addr; input: ReturnType<typeof coerceInputForCell> }[],
): void {
  const state = store.getState();
  const updates: { key: string; numFmt: NonNullable<CellFormat['numFmt']> }[] = [];
  for (const { addr, input } of writes) {
    if (input.kind !== 'number' || !input.implicitFormat) continue;
    const key = addrKey(addr);
    const current = state.format.formats.get(key)?.numFmt;
    if (current && current.kind !== 'general') continue;
    updates.push({ key, numFmt: input.implicitFormat });
  }
  if (updates.length === 0) return;
  recordFormatChange(history, store, () => {
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const { key, numFmt } of updates) formats.set(key, { ...formats.get(key), numFmt });
      return { ...s, format: { ...s.format, formats } };
    });
  });
}
