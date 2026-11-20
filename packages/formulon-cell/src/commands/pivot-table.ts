import type { PivotDefinition } from '../engine/pivot-definition.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { PivotFilterSpec, Range } from '../engine/types.js';
import {
  type PivotAggregation,
  PivotAxis,
  type PivotShowValuesAs,
  PivotSortOrder,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { sameRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import type { AutofitOptions } from './autofit-measurement.js';
import { currentRegion } from './current-region.js';
import { defaultObjectNames } from './format-as-table.js';
import { type History, recordingHistory, transactIn } from './history.js';
import type { OperationEffect } from './interaction-policy.js';
import {
  type PivotAuthorization,
  pivotFilterBlockRows,
  pivotReportRange,
  pivotWrite,
} from './pivot-actions.js';
import {
  type PivotSourceField,
  pivotCacheItemLabels,
  pivotCacheSourceFields,
  refreshFailure,
  writePivotCacheWorksheetSource,
} from './pivot-cache-source.js';
import { pivotFieldRoles, resolvePivotFieldPlan, valueFieldFormat } from './pivot-field-plan.js';
import { recreatePivotTable } from './pivot-recreate.js';
import { finishPivotReportUpdate, reloadSheetIfShown } from './pivot-report-update.js';
import { addSheet, moveSheet, removeSheet } from './sheet-mutate.js';

export type PivotSortDirection = 'none' | 'asc' | 'desc';

export interface PivotFieldItemVisibility {
  fieldName: string;
  itemName: string;
  visible: boolean;
}

export interface PivotValueFieldSettings {
  fieldName: string;
  aggregation?: PivotAggregation;
  numberFormat?: string;
  showValuesAs?: PivotShowValuesAs;
  /** Row or column field the show-as base axis refers to. */
  showAsBaseField?: string;
  /** Position among the base field's cache items, or a base-item sentinel. */
  showAsBaseItem?: number;
}

export interface CreatePivotTableOptions {
  source: Range;
  destination: { sheet: number; row: number; col: number };
  name?: string;
  /** Empty, with no value, column or filter field, creates an empty report
   *  whose fields are laid out afterwards from the field list. */
  rowField: string;
  columnField?: string;
  filterField?: string;
  filterFields?: readonly string[];
  filterItems?: readonly PivotFieldItemVisibility[];
  /** Label, value or date filters; each goes on the report's row or column field. */
  pivotFilters?: readonly PivotFilterSpec[];
  valueField: string;
  valueFields?: readonly string[];
  valueFieldSettings?: readonly PivotValueFieldSettings[];
  aggregation?: PivotAggregation;
  rowSort?: PivotSortDirection;
  columnSort?: PivotSortDirection;
  rowSubtotalTop?: boolean;
  columnSubtotalTop?: boolean;
  valueNumberFormat?: string;
  showRowGrandTotals?: boolean;
  showColumnGrandTotals?: boolean;
  /** An existing cache of the same source for the report to share; omitted,
   *  a new cache is built from the cells. */
  cacheId?: number;
}

export type CreatePivotTableResult =
  | { ok: true; cacheId: number; pivotIndex: number; sharedCache: boolean }
  | {
      ok: false;
      reason: 'unsupported' | 'invalid-range' | 'invalid-field' | 'engine-failed';
      step?: string;
    };

/** A source range read the way a new pivot cache reads it: the fields the engine names from the
 *  header row and each field's distinct items. Leaves no cache behind; null when the range cannot
 *  back a cache. */
export function readPivotSource(
  wb: WorkbookHandle,
  source: Range,
): { fields: PivotSourceField[]; items: string[][] } | null {
  const built = buildPivotCache(wb, source);
  if (typeof built !== 'number') return null;
  const fields = pivotCacheSourceFields(wb, built);
  const items = fields.map((field) => pivotCacheItemLabels(wb, built, field.index));
  wb.removePivotCache(built);
  return { fields, items };
}

/** The first free "<stem>N" report name, stem in the workbook's language;
 *  names compare case-insensitively. */
export const nextPivotTableName = (wb: WorkbookHandle): string => {
  const taken = new Set(
    wb
      .getPivotTables()
      .map((p) => wb.pivotTableName(p.sheetIndex, p.pivotIndex)?.toLocaleLowerCase()),
  );
  const { pivot } = defaultObjectNames(wb);
  let n = 1;
  while (taken.has(`${pivot}${n}`.toLocaleLowerCase())) n += 1;
  return `${pivot}${n}`;
};

/** A new cache the engine builds from `source`: its fields, shared items and records. */
function buildPivotCache(
  wb: WorkbookHandle,
  source: Range,
): number | Extract<CreatePivotTableResult, { ok: false }> {
  const cacheId = wb.createPivotCache(0);
  if (cacheId < 0) return { ok: false, reason: 'engine-failed', step: 'cache' };
  if (!writePivotCacheWorksheetSource(wb, cacheId, source)) {
    wb.removePivotCache(cacheId);
    return { ok: false, reason: 'engine-failed', step: 'cache-source' };
  }
  const refused = wb.refreshPivotCache(cacheId);
  if (refused !== null) {
    wb.removePivotCache(cacheId);
    return refreshFailure(refused);
  }
  return cacheId;
}

export function createPivotTableFromRange(
  wb: WorkbookHandle,
  opts: CreatePivotTableOptions,
): CreatePivotTableResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };

  const sharedCache = opts.cacheId !== undefined;
  const built = opts.cacheId ?? buildPivotCache(wb, opts.source);
  if (typeof built !== 'number') return built;
  const cacheId = built;
  const refuse = (reason: 'invalid-range' | 'invalid-field'): CreatePivotTableResult => {
    if (!sharedCache) wb.removePivotCache(cacheId);
    return { ok: false, reason };
  };
  const fields = pivotCacheSourceFields(wb, cacheId);
  const plan = resolvePivotFieldPlan(fields, opts);
  if (typeof plan === 'string') return refuse(plan);
  const { rowField, columnField, filterFields, valueFields, valueFieldSettings } = plan;

  const pivotIndex = wb.createPivotTable(
    opts.destination.sheet,
    opts.name?.trim() || nextPivotTableName(wb),
    cacheId,
    { row: opts.destination.row, col: opts.destination.col },
  );
  if (pivotIndex < 0) {
    if (!sharedCache) wb.removePivotCache(cacheId);
    return { ok: false, reason: 'engine-failed', step: 'pivot' };
  }
  const failAfterPivot = (step: string): CreatePivotTableResult => {
    wb.removePivotTable(opts.destination.sheet, pivotIndex);
    if (!sharedCache) wb.removePivotCache(cacheId);
    return { ok: false, reason: 'engine-failed', step };
  };

  if (
    !wb.setPivotTableGrandTotals(
      opts.destination.sheet,
      pivotIndex,
      opts.showRowGrandTotals ?? true,
      opts.showColumnGrandTotals ?? true,
    )
  ) {
    return failAfterPivot('grand-totals');
  }

  // Pivot field i binds to cache field i, so every cache field is added in
  // order; the ones the report does not use are then taken off their axis.
  const roles = pivotFieldRoles(plan);
  const roleSteps = {
    [PivotAxis.Row]: 'row-field',
    [PivotAxis.Col]: 'col-field',
    [PivotAxis.Page]: 'filter-field',
    [PivotAxis.Value]: 'value-field',
  } as const;
  const pivotFieldByName = new Map<string, number>();
  for (const field of fields) {
    const used = roles.get(field.name);
    const axis = used ?? PivotAxis.Value;
    const pivotField = wb.addPivotField(opts.destination.sheet, pivotIndex, {
      sourceName: field.name,
      axis,
      ...(axis === PivotAxis.Row ? { subtotalTop: opts.rowSubtotalTop ?? true } : {}),
      ...(axis === PivotAxis.Col ? { subtotalTop: opts.columnSubtotalTop ?? true } : {}),
    });
    if (pivotField < 0) return failAfterPivot(roleSteps[axis]);
    if (
      used === undefined &&
      !wb.setPivotFieldAxis(opts.destination.sheet, pivotIndex, pivotField, null)
    ) {
      return failAfterPivot(roleSteps[axis]);
    }
    pivotFieldByName.set(field.name, pivotField);
  }
  const rowPivotField = rowField ? (pivotFieldByName.get(rowField.name) as number) : -1;
  if (rowField && opts.rowSort && opts.rowSort !== 'none') {
    const ok = wb.setPivotFieldSort(
      opts.destination.sheet,
      pivotIndex,
      rowPivotField,
      opts.rowSort === 'asc' ? PivotSortOrder.Ascending : PivotSortOrder.Descending,
      '',
    );
    if (!ok) return failAfterPivot('row-sort');
  }

  const colPivotField = columnField ? (pivotFieldByName.get(columnField.name) as number) : -1;
  if (columnField && opts.columnSort && opts.columnSort !== 'none') {
    const ok = wb.setPivotFieldSort(
      opts.destination.sheet,
      pivotIndex,
      colPivotField,
      opts.columnSort === 'asc' ? PivotSortOrder.Ascending : PivotSortOrder.Descending,
      '',
    );
    if (!ok) return failAfterPivot('col-sort');
  }

  for (const filter of opts.pivotFilters ?? []) {
    if (!wb.addPivotFilter(opts.destination.sheet, pivotIndex, filter)) {
      return failAfterPivot('pivot-filter');
    }
  }

  for (const filterField of filterFields) {
    const filterPivotField = pivotFieldByName.get(filterField.name) as number;
    const filterItems = opts.filterItems?.filter((entry) => entry.fieldName === filterField.name);
    if (
      filterItems?.length &&
      !wb.clearPivotFieldItems(opts.destination.sheet, pivotIndex, filterPivotField)
    ) {
      return failAfterPivot('filter-item');
    }
    for (const item of filterItems ?? []) {
      if (
        !wb.addPivotFieldItem(
          opts.destination.sheet,
          pivotIndex,
          filterPivotField,
          item.itemName,
          item.visible,
        )
      ) {
        return failAfterPivot('filter-item');
      }
    }
  }

  // A field lists its items when it joins an axis, so a base item already has its position.
  const baseFieldIndexes = new Map<string, number>();
  if (rowField) baseFieldIndexes.set(rowField.name, rowPivotField);
  if (columnField) baseFieldIndexes.set(columnField.name, colPivotField);

  for (const valueField of valueFields) {
    const settings = valueFieldSettings.get(valueField.name);
    const { aggregation, numberFormat } = valueFieldFormat(valueField, settings, opts);

    // Left unnamed, the engine names the field from the workbook profile.
    const dataField = wb.addPivotDataField(opts.destination.sheet, pivotIndex, {
      fieldIndex: pivotFieldByName.get(valueField.name) as number,
      aggregation,
      numberFormat: numberFormat || undefined,
      showValuesAs: settings?.showValuesAs,
      showAsBaseField:
        settings?.showAsBaseField === undefined
          ? undefined
          : baseFieldIndexes.get(settings.showAsBaseField),
      showAsBaseItem: settings?.showAsBaseItem,
    });
    if (dataField < 0) return failAfterPivot('data-field');
  }

  return { ok: true, cacheId, pivotIndex, sharedCache };
}

/** Where a new report lands: a fresh sheet inserted before the source sheet
 *  (report at A3), or a cell of the source sheet. */
export type PivotPlacement = 'new-sheet' | { row: number; col: number };

export const PIVOT_NEW_SHEET_ANCHOR = { row: 2, col: 0 } as const;

export type CreatePlacedPivotResult =
  | {
      ok: true;
      cacheId: number;
      pivotIndex: number;
      destination: { sheet: number; row: number; col: number };
    }
  | {
      ok: false;
      reason:
        | 'refused'
        | 'sheet-failed'
        | 'unsupported'
        | 'invalid-range'
        | 'invalid-field'
        | 'engine-failed';
      step?: string;
    };

export interface CreatePlacedPivotDeps extends PivotAuthorization {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history?: History | null;
  /** Measures the report columns refitted when an undo or redo brings the report back. */
  autofit?: AutofitOptions;
}

/** Source range for a new report: the selection, or for a single cell the
 *  Table containing it or else its current region. */
export function pivotSourceForSelection(store: SpreadsheetStore, wb: WorkbookHandle): Range {
  const state = store.getState();
  const { range } = state.selection;
  if (range.r0 !== range.r1 || range.c0 !== range.c1) return range;
  for (const table of wb.getTables()) {
    if (table.sheetIndex !== range.sheet) continue;
    const parsed = parseRangeRef(table.ref);
    if (
      parsed &&
      range.r0 >= parsed.r0 &&
      range.r0 <= parsed.r1 &&
      range.c0 >= parsed.c0 &&
      range.c0 <= parsed.c1
    ) {
      return { sheet: range.sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 };
    }
  }
  return currentRegion(state, { sheet: range.sheet, row: range.r0, col: range.c0 }) ?? range;
}

/** Create a report at `placement` as one authorized undo step (the new sheet,
 *  its position and the report revert together). Its effect is the anchor cell
 *  on an existing sheet, or a new sheet. */
export function createPivotTableAtPlacement(
  deps: CreatePlacedPivotDeps,
  opts: Omit<CreatePivotTableOptions, 'destination'>,
  placement: PivotPlacement,
): CreatePlacedPivotResult {
  const newSheet = placement === 'new-sheet';
  const effects: readonly OperationEffect[] = newSheet
    ? [{ kind: 'workbook' }]
    : [{ kind: 'cells', cells: [{ sheet: opts.source.sheet, ...placement }] }];
  return pivotWrite(
    deps,
    deps.history,
    effects,
    { ok: false, reason: 'refused' },
    () => createPlacedReport(deps, opts, placement),
    { addsSheet: newSheet },
  );
}

function createPlacedReport(
  deps: CreatePlacedPivotDeps,
  opts: Omit<CreatePivotTableOptions, 'destination'>,
  placement: PivotPlacement,
): CreatePlacedPivotResult {
  const { store, workbook: wb } = deps;
  const history = recordingHistory(deps.history);
  return transactIn(
    history,
    (): CreatePlacedPivotResult => {
      let source = opts.source;
      let destination = {
        sheet: source.sheet,
        ...(placement === 'new-sheet' ? PIVOT_NEW_SHEET_ANCHOR : placement),
      };
      let addedSheet = -1;
      if (placement === 'new-sheet') {
        addedSheet = addSheet(store, wb, history);
        if (addedSheet < 0) return { ok: false, reason: 'sheet-failed' };
        // The new sheet takes the source sheet's slot, pushing the source right.
        if (moveSheet(store, wb, addedSheet, source.sheet, history)) {
          destination = { ...destination, sheet: source.sheet };
          source = { ...source, sheet: source.sheet + 1 };
        } else {
          destination = { ...destination, sheet: addedSheet };
        }
      }
      const create = (): CreatePivotTableResult =>
        createPivotTableFromRange(wb, { ...opts, source, destination });
      const result = create();
      if (!result.ok) {
        // A history rolls the added sheet back with the rest of the step.
        if (!history && addedSheet >= 0) removeSheet(store, wb, addedSheet);
        return result;
      }
      if (history) {
        let current = result;
        const refresh = (): void => {
          if (store.getState().data.sheetIndex === destination.sheet) {
            finishPivotReportUpdate(store, wb, deps.autofit);
          }
        };
        history.push({
          undo: () => {
            wb.removePivotTable(destination.sheet, current.pivotIndex);
            if (!current.sharedCache) wb.removePivotCache(current.cacheId);
            refresh();
          },
          redo: () => {
            const again = create();
            if (again.ok) current = again;
            refresh();
          },
        });
      }
      return { ok: true, cacheId: result.cacheId, pivotIndex: result.pivotIndex, destination };
    },
    { keep: (placed) => placed.ok },
  );
}

/** The report whose whole extent is exactly `range`, or null. */
export function pivotTableAtRange(wb: WorkbookHandle, range: Range): number | null {
  if (!wb.capabilities.pivotTables) return null;
  const hit = wb.pivotBounds(range.sheet).find((b) => sameRange(b.range, range));
  return hit ? hit.pivotIndex : null;
}

/** Paste a whole report as a second report over the same cache with its
 *  top-left at `at`, named as a new report is; one undo step. Returns the
 *  rectangle it occupies, or null when it cannot be rebuilt. */
export function pastePivotTableCopy(
  deps: CreatePlacedPivotDeps,
  source: { sheet: number; pivotIndex: number },
  at: { sheet: number; row: number; col: number },
): Range | null {
  return pivotWrite(
    deps,
    deps.history,
    [{ kind: 'cells', cells: [{ sheet: at.sheet, row: at.row, col: at.col }] }],
    null,
    () => pasteReportCopy(deps, source, at),
  );
}

function pasteReportCopy(
  deps: CreatePlacedPivotDeps,
  source: { sheet: number; pivotIndex: number },
  at: { sheet: number; row: number; col: number },
): Range | null {
  const { store, workbook: wb } = deps;
  const history = recordingHistory(deps.history);
  const definition = wb.getPivotDefinition(source.sheet, source.pivotIndex);
  if (!definition) return null;
  // The extent starts at the filter block, one blank row above the anchor.
  const anchor = { row: at.row + pivotFilterBlockRows(definition), col: at.col };
  const copy = { ...definition, name: nextPivotTableName(wb) };
  let pivotIndex = recreatePivotTable(wb, at.sheet, anchor, copy);
  if (pivotIndex < 0) return null;
  const refresh = (): void => reloadSheetIfShown(store, wb, at.sheet);
  history?.push({
    undo: () => {
      wb.removePivotTable(at.sheet, pivotIndex);
      refresh();
    },
    redo: () => {
      pivotIndex = recreatePivotTable(wb, at.sheet, anchor, copy);
      refresh();
    },
  });
  return pivotReportRange(wb, at.sheet, pivotIndex);
}

export interface DeletePivotTablesDeps extends PivotAuthorization {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history?: History | null;
}

/** Remove whole reports as one undo step. Undo rebuilds each report from the
 *  definition read back before removal; the pivot cache stays in the
 *  workbook meanwhile. False, with nothing removed, when a definition cannot
 *  be read back (so the removal could not be undone). */
export function deletePivotTables(
  deps: DeletePivotTablesDeps,
  sheet: number,
  pivots: readonly { pivotIndex: number; range: Range }[],
): boolean {
  return pivotWrite(
    deps,
    deps.history,
    pivots.map(({ range }) => ({ kind: 'range' as const, range: { ...range } })),
    false,
    () => removeReports(deps, sheet, pivots),
  );
}

function removeReports(
  deps: DeletePivotTablesDeps,
  sheet: number,
  pivots: readonly { pivotIndex: number; range: Range }[],
): boolean {
  const { store, workbook: wb } = deps;
  const history = recordingHistory(deps.history);
  const ascending = [...pivots].sort((a, b) => a.pivotIndex - b.pivotIndex);
  const captured: { definition: PivotDefinition; anchor: { row: number; col: number } }[] = [];
  for (const { pivotIndex, range } of ascending) {
    const definition = wb.getPivotDefinition(sheet, pivotIndex);
    if (!definition) return false;
    // The extent starts at the filter block, which sits above the anchor with
    // one blank row between them.
    captured.push({
      definition,
      anchor: { row: range.r0 + pivotFilterBlockRows(definition), col: range.c0 },
    });
  }
  const refresh = (): void => reloadSheetIfShown(store, wb, sheet);
  // Highest index first so the lower ones stay valid while removing.
  const removeAll = (indices: readonly number[]): void => {
    for (const index of [...indices].sort((a, b) => b - a)) wb.removePivotTable(sheet, index);
    refresh();
  };
  removeAll(ascending.map((pivot) => pivot.pivotIndex));
  if (history) {
    let restored: number[] = [];
    history.push({
      undo: () => {
        restored = captured
          .map(({ definition, anchor }) => recreatePivotTable(wb, sheet, anchor, definition))
          .filter((index) => index >= 0);
        refresh();
      },
      redo: () => removeAll(restored),
    });
  }
  return true;
}
