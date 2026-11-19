import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import type { Range } from '../engine/types.js';
import { PivotAggregation } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { sameRange } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { pivotReportEffects, pivotWrite } from './pivot-actions.js';
import { pivotCacheWorksheetSourceRange, refreshPivotTable } from './pivot-cache-source.js';
import { refreshPivotReportsAfterEdit } from './pivot-report-update.js';
import {
  createPivotTableAtPlacement,
  nextPivotTableName,
  pivotSourceForSelection,
  readPivotSource,
} from './pivot-table.js';

export type RibbonPivotTableAction =
  | 'dialog'
  | 'recommended'
  | 'recommended-insert'
  | 'new-sheet'
  | 'existing-sheet'
  | 'refresh';

export interface RibbonPivotTableReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

export interface RibbonPivotTableReport {
  title: string;
  items: RibbonPivotTableReportItem[];
}

export type RibbonPivotTableActionResult =
  | { kind: 'open-dialog' }
  | {
      kind: 'created';
      destinationSheet: number;
      destination: { sheet: number; row: number; col: number };
    }
  | { kind: 'refreshed'; sheet: number }
  | { kind: 'report'; report: RibbonPivotTableReport };

export interface RibbonPivotTableActionStrings {
  pivotTable: string;
  pivotTableNewSheet: string;
  pivotTableRefreshData: string;
  pivotTableRefreshUnavailable: string;
  recommendedPivotTables: string;
  pivotAuthoringDetail: string;
  workbookStructureProtectedBlocked: string;
}

export interface ExecuteRibbonPivotTableActionDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  /** The entry point the action runs under. */
  origin: InteractionOrigin;
  commandId: string;
  action: RibbonPivotTableAction;
  strings: RibbonPivotTableActionStrings;
  history?: History | null;
}

/** A one-line report dialog result titled `title`. */
const pivotReport = (
  strings: RibbonPivotTableActionStrings,
  title: string,
  severity: RibbonPivotTableReportItem['severity'],
  detail: string,
): RibbonPivotTableActionResult => ({
  kind: 'report',
  report: { title, items: [{ severity, label: strings.pivotTable, detail }] },
});

/** Show the report's new sheet, settle its reports and select its first cell. */
function createdResult(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  destination: { sheet: number; row: number; col: number },
): RibbonPivotTableActionResult {
  mutators.setSheetIndex(store, destination.sheet);
  refreshPivotReportsAfterEdit(store, workbook);
  mutators.setActive(store, destination);
  return { kind: 'created', destinationSheet: destination.sheet, destination };
}

/** An existing cache whose worksheet source is exactly `source`. */
const cacheForSource = (wb: WorkbookHandle, source: Range): number | undefined =>
  wb.pivotCacheIds().find((id) => {
    const range = pivotCacheWorksheetSourceRange(wb, id);
    return range !== null && sameRange(range, source);
  });

/**
 * Recommended PivotTables as the Mac desktop runs it: no dialog, a report on a
 * new sheet before the source sheet at A3. Measured on one fixture only: the
 * first text field goes to rows, the first numeric field to columns, and there
 * are no values. A cache that already holds the source is shared, as there.
 */
function insertRecommendedPivotTable(
  deps: ExecuteRibbonPivotTableActionDeps,
): RibbonPivotTableActionResult {
  const { store, workbook, strings, history = null } = deps;
  const source = pivotSourceForSelection(store, workbook);
  const fields = readPivotSource(workbook, source)?.fields ?? [];
  const rowField = fields.find((field) => !field.numeric);
  const columnField = fields.find((field) => field.numeric);
  if (!workbook.capabilities.pivotTableMutate || !rowField || !columnField) {
    return pivotReport(
      strings,
      strings.recommendedPivotTables,
      'warning',
      strings.pivotAuthoringDetail,
    );
  }
  const cacheId = cacheForSource(workbook, source);
  const result = createPivotTableAtPlacement(
    { store, workbook, history, origin: deps.origin, commandId: deps.commandId },
    {
      source,
      name: nextPivotTableName(workbook),
      rowField: rowField.name,
      columnField: columnField.name,
      valueField: '',
      ...(cacheId === undefined ? {} : { cacheId }),
    },
    'new-sheet',
  );
  if (!result.ok) {
    return pivotReport(
      strings,
      strings.recommendedPivotTables,
      'warning',
      result.reason === 'sheet-failed'
        ? strings.workbookStructureProtectedBlocked
        : strings.pivotAuthoringDetail,
    );
  }
  return createdResult(store, workbook, result.destination);
}

/** Implements the cross-host "PivotTable" ribbon split-button. Returns one of:
 *  - `open-dialog` — host should run its own `openPivotTableDialog` flow.
 *  - `created` — engine wrote the pivot; `mutators.replaceCells / setSheetIndex /
 *    setActive` have already been applied, the host only needs to mirror its
 *    local active-state cache.
 *  - `report` — show this report dialog to the user (unsupported workbook,
 *    insufficient source data, Recommended PivotTables, blocked by structure
 *    protection, or engine refusal to author the pivot).
 *
 *  Encapsulating the branching here keeps React, Vue, and the playground
 *  shells identical and removes ~100 lines of duplication per host. */
export const executeRibbonPivotTableAction = (
  deps: ExecuteRibbonPivotTableActionDeps,
): RibbonPivotTableActionResult => {
  const { store, workbook, action, strings, history = null } = deps;
  if (action === 'refresh') {
    const active = store.getState().selection.active;
    const pivot = findPivotTableAtCell(workbook, active);
    const refreshed = pivot
      ? pivotWrite(
          deps,
          history,
          pivotReportEffects(workbook, pivot.sheetIndex, pivot.pivotIndex),
          { ok: false } as ReturnType<typeof refreshPivotTable>,
          () =>
            refreshPivotTable(workbook, { sheet: pivot.sheetIndex, pivotIndex: pivot.pivotIndex }),
        ).ok
      : false;
    if (!pivot || !refreshed) {
      return pivotReport(
        strings,
        strings.pivotTableRefreshData,
        'warning',
        strings.pivotTableRefreshUnavailable,
      );
    }
    mutators.replaceCells(store, workbook.cells(pivot.sheetIndex));
    return { kind: 'refreshed', sheet: pivot.sheetIndex };
  }
  if (action === 'dialog' || action === 'existing-sheet') {
    return { kind: 'open-dialog' };
  }
  if (action === 'recommended-insert') return insertRecommendedPivotTable(deps);
  if (action === 'recommended') {
    return pivotReport(
      strings,
      strings.recommendedPivotTables,
      'info',
      strings.pivotAuthoringDetail,
    );
  }
  if (!workbook.capabilities.pivotTableMutate) {
    return pivotReport(strings, strings.pivotTableNewSheet, 'info', strings.pivotAuthoringDetail);
  }
  const source = pivotSourceForSelection(store, workbook);
  const fields = readPivotSource(workbook, source)?.fields ?? [];
  const valueField = fields.find((field) => field.numeric) ?? fields.at(-1);
  const rowField = fields.find((field) => field.name !== valueField?.name) ?? fields[0];
  if (!rowField || !valueField || rowField.name === valueField.name) {
    return pivotReport(strings, strings.pivotTable, 'warning', strings.pivotAuthoringDetail);
  }
  const result = createPivotTableAtPlacement(
    { store, workbook, history, origin: deps.origin, commandId: deps.commandId },
    {
      source,
      name: nextPivotTableName(workbook),
      rowField: rowField.name,
      valueField: valueField.name,
      aggregation: valueField.numeric ? PivotAggregation.Sum : PivotAggregation.Count,
    },
    'new-sheet',
  );
  if (result.ok) return createdResult(store, workbook, result.destination);
  if (result.reason === 'sheet-failed') {
    return pivotReport(
      strings,
      strings.pivotTableNewSheet,
      'warning',
      strings.workbookStructureProtectedBlocked,
    );
  }
  return pivotReport(strings, strings.pivotTableNewSheet, 'info', strings.pivotAuthoringDetail);
};
