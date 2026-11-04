import type { History } from '../commands/history.js';
import {
  PIVOT_NEW_SHEET_ANCHOR,
  recreatePivotTable,
  refreshPivotReportsAfterEdit,
} from '../commands/pivot-table.js';
import { addSheet, moveSheet } from '../commands/sheet-mutate.js';
import { PivotAxis, type Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { commitPivotChange } from './pivot-field-settings-dialog-apply.js';

export interface MovePivotDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history?: History | null;
}

/** Where a report goes: the top-left of its extent on a sheet, or a new sheet
 *  placed right after the report's own. */
export type PivotMoveTarget = { sheet: number; row: number; col: number } | 'new-sheet';

/** Move a whole report as one undo step. The report's extent keeps its size;
 *  `target` is where its top-left lands. Returns the extent it occupies
 *  afterwards, or null when nothing moved. */
export function movePivotTable(
  deps: MovePivotDeps,
  source: { sheet: number; pivotIndex: number },
  target: PivotMoveTarget,
): Range | null {
  const { store, workbook: wb } = deps;
  const history = deps.history && !deps.history.isReplaying() ? deps.history : null;
  const definition = wb.getPivotDefinition(source.sheet, source.pivotIndex);
  const bounds = wb
    .pivotBounds(source.sheet)
    .find((b) => b.pivotIndex === source.pivotIndex)?.range;
  if (!definition || !bounds) return null;
  const filterRows = definition.fields.filter((f) => f.axis === PivotAxis.Page).length;
  const anchorOffset = filterRows > 0 ? filterRows + 1 : 0;
  const rows = bounds.r1 - bounds.r0 + 1;
  const cols = bounds.c1 - bounds.c0 + 1;

  if (target !== 'new-sheet' && target.sheet === source.sheet) {
    const ctx = {
      wb,
      history,
      onChanged: () => {
        refreshPivotReportsAfterEdit(store, wb);
      },
    };
    const moved = commitPivotChange(ctx, source.sheet, source.pivotIndex, () =>
      wb.setPivotTableAnchor(source.sheet, source.pivotIndex, {
        row: target.row + anchorOffset,
        col: target.col,
        rows: rows - anchorOffset,
        cols,
      }),
    );
    return moved ? selectExtent(deps, source.sheet, source.pivotIndex) : null;
  }

  const txn = history?.begin();
  let sheet: number;
  let at: { row: number; col: number };
  if (target === 'new-sheet') {
    const added = addSheet(store, wb, history);
    if (added < 0) {
      if (history && txn) history.end(txn);
      return null;
    }
    // The new sheet sits right after the report's sheet.
    sheet = added;
    const slot = source.sheet + 1;
    if (added !== slot && moveSheet(store, wb, added, slot, history)) sheet = slot;
    at = PIVOT_NEW_SHEET_ANCHOR;
  } else {
    sheet = target.sheet;
    at = target;
  }
  const anchor = { row: at.row + anchorOffset, col: at.col };
  const refresh = (...sheets: number[]): void => {
    for (const s of sheets) {
      if (store.getState().data.sheetIndex === s) {
        mutators.replaceCells(store, wb.cells(s));
      }
    }
  };
  const oldAnchor = { row: bounds.r0 + anchorOffset, col: bounds.c0 };
  wb.removePivotTable(source.sheet, source.pivotIndex);
  let created = recreatePivotTable(wb, sheet, anchor, definition);
  if (created < 0) {
    recreatePivotTable(wb, source.sheet, oldAnchor, definition);
    refresh(source.sheet);
    if (history && txn) history.abort(txn);
    return null;
  }
  let removed = source.pivotIndex;
  history?.push({
    undo: () => {
      wb.removePivotTable(sheet, created);
      removed = recreatePivotTable(wb, source.sheet, oldAnchor, definition);
      refresh(sheet, source.sheet);
    },
    redo: () => {
      wb.removePivotTable(source.sheet, removed);
      created = recreatePivotTable(wb, sheet, anchor, definition);
      refresh(sheet, source.sheet);
    },
  });
  if (history && txn) history.end(txn);
  mutators.setSheetIndex(store, sheet);
  refreshPivotReportsAfterEdit(store, wb);
  return selectExtent(deps, sheet, created);
}

/** Select the report's whole extent with its top-left active. */
function selectExtent(deps: MovePivotDeps, sheet: number, pivotIndex: number): Range | null {
  const { store, workbook: wb } = deps;
  const range = wb.pivotBounds(sheet).find((b) => b.pivotIndex === pivotIndex)?.range;
  if (!range) return null;
  if (store.getState().data.sheetIndex !== sheet) mutators.setSheetIndex(store, sheet);
  mutators.setActive(store, { sheet, row: range.r0, col: range.c0 });
  mutators.setRange(store, range);
  return range;
}
