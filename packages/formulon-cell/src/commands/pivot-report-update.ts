import type { PivotSpill } from '../engine/pivot-definition.js';
import { lastPivotStyleLayer, pivotStyleLayer } from '../engine/pivot-style.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { AutofitOptions } from './autofit-measurement.js';
import { pivotFilterBlockRows, pivotReportRange } from './pivot-actions.js';
import { reflowPivotFormats } from './pivot-format-reflow.js';
import { autofitColsWidthForRanges } from './row-col-layout.js';

/** Fit the columns of the PivotTables on the sheet on show to their content,
 *  as the desktop does after a report update; a column can shrink, and columns
 *  outside the reports keep their width. A report whose "autofit column widths
 *  on update" option is off keeps its columns. Part of the update, so no undo step. */
export function autofitPivotTableColumns(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  const state = store.getState();
  // An empty report keeps the columns it is drawn over.
  const empty = new Set(pivotStyleLayer(wb, state).placeholders.map((p) => p.pivotIndex));
  const ranges = wb
    .pivotBounds(state.data.sheetIndex)
    .filter((b) => !empty.has(b.pivotIndex) && wb.pivotAutofitColumns(b.range.sheet, b.pivotIndex))
    .map((b) => b.range);
  if (ranges.length > 0) autofitColsWidthForRanges(store, null, ranges, wb, opts);
}

/** Complete a report create or update on the sheet on show: reload its cells,
 *  then refit the report columns that opt in. */
export function finishPivotReportUpdate(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  mutators.refreshActiveSheetCells(store, wb);
  autofitPivotTableColumns(store, wb, opts);
}

/** Pin a report's anchor to the top of the report proper. A report filter
 *  block pushes a report with no room above it down; the report then stays
 *  there when the filter goes, as on the desktop. */
export function settlePivotAnchor(wb: WorkbookHandle, sheet: number, pivotIndex: number): void {
  const bounds = pivotReportRange(wb, sheet, pivotIndex);
  const definition = bounds ? wb.getPivotDefinition(sheet, pivotIndex) : null;
  if (!bounds || !definition) return;
  const filterBlock = pivotFilterBlockRows(definition);
  if (filterBlock === 0) return;
  const row = bounds.r0 + filterBlock;
  // A report held at #SPILL! shows only its anchor; leave it where it is.
  if (row > bounds.r1) return;
  wb.setPivotTableAnchor(sheet, pivotIndex, {
    row,
    col: bounds.c0,
    rows: bounds.r1 - row + 1,
    cols: bounds.c1 - bounds.c0 + 1,
  });
}

/** Blank the sheet's own cells inside the extents reports occupied before an
 *  edit wherever the reports no longer show a value: a report owns its extent,
 *  so values cached there from an earlier layout must not show through. */
function clearVacatedReportCells(
  wb: WorkbookHandle,
  sheet: number,
  extents: readonly Range[],
): void {
  if (extents.length === 0) return;
  const shown = new Map<string, string>();
  for (const cell of wb.pivotCells(sheet)) {
    shown.set(`${cell.addr.row}:${cell.addr.col}`, JSON.stringify(cell.value));
  }
  const stale: { row: number; col: number }[] = [];
  for (const cell of wb.physicalCells(sheet)) {
    const { row, col } = cell.addr;
    // A cached value matching what the report shows there stays; any other
    // shows through wherever the grid reads the sheet's own cell.
    if (cell.formula === null && shown.get(`${row}:${col}`) === JSON.stringify(cell.value))
      continue;
    if (extents.some((e) => row >= e.r0 && row <= e.r1 && col >= e.c0 && col <= e.c1)) {
      stale.push({ row, col });
    }
  }
  if (stale.length === 0) return;
  wb.clearPivotCachedCells(sheet, stale);
  wb.recalc();
}

/** Reload the projected cells of `sheet` when it is the shown sheet. */
export function reloadSheetIfShown(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
): void {
  if (store.getState().data.sheetIndex === sheet) mutators.replaceCells(store, wb.cells(sheet));
}

/** Bring the active sheet's reports up to date after a report edit, an undo
 *  or a redo: settle their anchors, reload the projected cells, move direct
 *  formats with the items they belong to and refit the report columns.
 *  Returns the reports held at #SPILL! by cells in their way. */
export function refreshPivotReportsAfterEdit(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  opts?: AutofitOptions,
): { spilled: PivotSpill[] } {
  const sheet = store.getState().data.sheetIndex;
  const before = lastPivotStyleLayer(wb, sheet);
  for (const { pivotIndex } of wb.pivotBounds(sheet)) settlePivotAnchor(wb, sheet, pivotIndex);
  clearVacatedReportCells(wb, sheet, before?.reports.map((r) => r.range) ?? []);
  mutators.replaceCells(store, wb.cells(sheet));
  reflowPivotFormats(store, wb);
  autofitPivotTableColumns(store, wb, opts);
  return { spilled: wb.pivotSpills(sheet) };
}
