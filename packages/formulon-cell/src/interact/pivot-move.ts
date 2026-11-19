import { type History, recordingHistory, transactIn } from '../commands/history.js';
import type { OperationEffect } from '../commands/interaction-policy.js';
import {
  type PivotAuthorization,
  pivotFilterBlockRows,
  pivotReportRange,
  pivotWrite,
} from '../commands/pivot-actions.js';
import { recreatePivotTable } from '../commands/pivot-recreate.js';
import {
  refreshPivotReportsAfterEdit,
  reloadSheetIfShown,
} from '../commands/pivot-report-update.js';
import { PIVOT_NEW_SHEET_ANCHOR } from '../commands/pivot-table.js';
import { addSheet, moveSheet } from '../commands/sheet-mutate.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { commitPivotChange } from './pivot-field-settings-dialog-apply.js';

export interface MovePivotDeps extends PivotAuthorization {
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
  const { workbook: wb } = deps;
  const bounds = pivotReportRange(wb, source.sheet, source.pivotIndex);
  if (!bounds) return null;
  const effects: OperationEffect[] = [{ kind: 'range', range: bounds }];
  if (target !== 'new-sheet') {
    effects.push({
      kind: 'range',
      range: {
        sheet: target.sheet,
        r0: target.row,
        c0: target.col,
        r1: Math.min(MAX_ROW, target.row + bounds.r1 - bounds.r0),
        c1: Math.min(MAX_COL, target.col + bounds.c1 - bounds.c0),
      },
    });
  }
  return pivotWrite(
    deps,
    deps.history,
    effects,
    null,
    () => moveReport(deps, source, target, bounds),
    {
      addsSheet: target === 'new-sheet',
    },
  );
}

function moveReport(
  deps: MovePivotDeps,
  source: { sheet: number; pivotIndex: number },
  target: PivotMoveTarget,
  bounds: Range,
): Range | null {
  const { store, workbook: wb } = deps;
  const history = recordingHistory(deps.history);
  const definition = wb.getPivotDefinition(source.sheet, source.pivotIndex);
  if (!definition) return null;
  const anchorOffset = pivotFilterBlockRows(definition);
  const rows = bounds.r1 - bounds.r0 + 1;
  const cols = bounds.c1 - bounds.c0 + 1;

  if (target !== 'new-sheet' && target.sheet === source.sheet) {
    const ctx = {
      wb,
      history,
      store: deps.store,
      origin: deps.origin,
      commandId: deps.commandId,
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

  const placed = transactIn(
    history,
    (): { sheet: number; created: number } | null => {
      let sheet: number;
      let at: { row: number; col: number };
      if (target === 'new-sheet') {
        const added = addSheet(store, wb, history);
        if (added < 0) return null;
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
        for (const s of sheets) reloadSheetIfShown(store, wb, s);
      };
      const oldAnchor = { row: bounds.r0 + anchorOffset, col: bounds.c0 };
      wb.removePivotTable(source.sheet, source.pivotIndex);
      let created = recreatePivotTable(wb, sheet, anchor, definition);
      if (created < 0) {
        recreatePivotTable(wb, source.sheet, oldAnchor, definition);
        refresh(source.sheet);
        return null;
      }
      const placedSheet = sheet;
      const placedIndex = created;
      let removed = source.pivotIndex;
      history?.push({
        undo: () => {
          wb.removePivotTable(placedSheet, created);
          removed = recreatePivotTable(wb, source.sheet, oldAnchor, definition);
          refresh(placedSheet, source.sheet);
        },
        redo: () => {
          wb.removePivotTable(source.sheet, removed);
          created = recreatePivotTable(wb, placedSheet, anchor, definition);
          refresh(placedSheet, source.sheet);
        },
      });
      return { sheet: placedSheet, created: placedIndex };
    },
    { keep: (result) => result !== null },
  );
  if (!placed) return null;
  const { sheet, created } = placed;
  mutators.setSheetIndex(store, sheet);
  refreshPivotReportsAfterEdit(store, wb);
  return selectExtent(deps, sheet, created);
}

/** Select the report's whole extent with its top-left active. */
function selectExtent(deps: MovePivotDeps, sheet: number, pivotIndex: number): Range | null {
  const { store, workbook: wb } = deps;
  const range = pivotReportRange(wb, sheet, pivotIndex);
  if (!range) return null;
  if (store.getState().data.sheetIndex !== sheet) mutators.setSheetIndex(store, sheet);
  mutators.setActive(store, { sheet, row: range.r0, col: range.c0 });
  mutators.setRange(store, range);
  return range;
}
