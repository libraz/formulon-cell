import { StatusCode } from '@libraz/formulon';
import { rangeOfRect, rectOfRange } from '../engine/address.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import { lastPivotStyleLayer, pivotStyleLayer } from '../engine/pivot-style.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { CellValue, PivotFilterSpec, Range } from '../engine/types.js';
import { PivotAggregation, PivotAxis, PivotShowValuesAs, PivotSortOrder } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { PivotDefinition, PivotSpill } from '../engine/workbook-handle-pivot.js';
import { sameRange } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { AutofitOptions } from './autofit-measurement.js';
import { currentRegion } from './current-region.js';
import { defaultObjectNames } from './format-as-table.js';
import { type History, recordingHistory, transactIn } from './history.js';
import type { InteractionOrigin, OperationEffect } from './interaction-policy.js';
import {
  type PivotAuthorization,
  pivotFilterBlockRows,
  pivotReportEffects,
  pivotReportRange,
  pivotWrite,
} from './pivot-actions.js';
import { reflowPivotFormats } from './pivot-format-reflow.js';
import { autofitColsWidthForRanges } from './row-col-layout.js';
import { addSheet, moveSheet, removeSheet } from './sheet-mutate.js';

export interface PivotSourceField {
  name: string;
  index: number;
  /** Whether the field's source column holds a number. */
  numeric: boolean;
}

export type PivotSortDirection = 'none' | 'asc' | 'desc';

export interface PivotFieldItemVisibility {
  fieldName: string;
  itemName: string;
  visible: boolean;
}

/** Show-as kinds that yield a ratio; the desktop app formats them "0.00%". */
const PERCENT_SHOW_AS: ReadonlySet<PivotShowValuesAs> = new Set([
  PivotShowValuesAs.PercentOfRow,
  PivotShowValuesAs.PercentOfCol,
  PivotShowValuesAs.PercentOfTotal,
  PivotShowValuesAs.PercentOf,
  PivotShowValuesAs.PercentDifferenceFrom,
  PivotShowValuesAs.PercentOfParentRow,
  PivotShowValuesAs.PercentOfParentCol,
  PivotShowValuesAs.PercentOfParent,
  PivotShowValuesAs.PercentOfRunningTotal,
]);
const PERCENT_FORMAT = '0.00%';

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

export interface RefreshPivotCacheOptions {
  cacheId: number;
  source: Range;
}

export interface RefreshPivotTableOptions {
  sheet: number;
  pivotIndex: number;
  source: Range;
}

export interface RefreshPivotTableFromStoredSourceOptions {
  sheet: number;
  pivotIndex: number;
}

export type RefreshPivotCacheResult =
  | { ok: true; cacheId: number }
  | {
      ok: false;
      reason: 'unsupported' | 'invalid-range' | 'invalid-field' | 'invalid-pivot' | 'engine-failed';
      step?: string;
    };

const cellLabel = (value: CellValue, fallback: string): string => {
  if (value.kind === 'text') return value.value.trim() || fallback;
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  return fallback;
};

const writePivotCacheWorksheetSource = (
  wb: WorkbookHandle,
  cacheId: number,
  source: Range,
): boolean => {
  if (!wb.capabilities.pivotCacheSource) return true;
  return wb.setPivotCacheWorksheetSource(cacheId, {
    range: {
      sheet: source.sheet,
      rect: rectOfRange(source),
    },
  });
};

export const pivotCacheWorksheetSourceRange = (
  wb: WorkbookHandle,
  cacheId: number,
): Range | null => {
  const range = wb.getPivotCacheWorksheetSource(cacheId)?.range;
  if (!range) return null;
  return rangeOfRect(range.sheet, range.rect);
};

/** The fields of a pivot cache, as the engine named and typed them from the source's header row. */
export const pivotCacheSourceFields = (wb: WorkbookHandle, cacheId: number): PivotSourceField[] =>
  wb
    .pivotCacheFields(cacheId)
    .map((field, index) => ({ name: field.name, index, numeric: field.containsNumber }));

/** Whether the source column behind a cache field holds numbers. */
export function pivotCacheFieldIsNumeric(
  wb: WorkbookHandle,
  cacheId: number,
  cacheFieldIndex: number,
): boolean {
  return wb.pivotCacheFields(cacheId)[cacheFieldIndex]?.containsNumber ?? false;
}

/** The distinct items of a cache field, in cache order, as labels. */
export const pivotCacheItemLabels = (
  wb: WorkbookHandle,
  cacheId: number,
  cacheFieldIndex: number,
): string[] =>
  wb.pivotCacheSharedItems(cacheId, cacheFieldIndex).map((v) => cellLabel(v, '(blank)'));

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

/** A refusal of the engine's Refresh: a source that does not resolve or is too large is the
 *  range's fault, anything else the engine's. */
const refreshFailure = (
  code: number,
): { ok: false; reason: 'invalid-range' | 'engine-failed'; step: string } => ({
  ok: false,
  reason:
    code === StatusCode.PreconditionFailed || code === StatusCode.SecurityResourceLimit
      ? 'invalid-range'
      : 'engine-failed',
  step: 'refresh',
});

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
  if (fields.length < 2) return refuse('invalid-range');
  const byName = new Map(fields.map((f) => [f.name, f]));
  const rowField = byName.get(opts.rowField);
  const columnField = opts.columnField ? byName.get(opts.columnField) : undefined;
  const filterFieldNames = Array.from(
    new Set(
      (opts.filterFields && opts.filterFields.length > 0
        ? opts.filterFields
        : opts.filterField
          ? [opts.filterField]
          : []
      ).filter(Boolean),
    ),
  );
  const filterFields = filterFieldNames.map((name) => byName.get(name));
  const pivotFilterFields = new Set((opts.pivotFilters ?? []).map((filter) => filter.fieldName));
  const valueFieldNames = Array.from(
    new Set(
      (opts.valueFields && opts.valueFields.length > 0
        ? opts.valueFields
        : [opts.valueField]
      ).filter(Boolean),
    ),
  );
  const valueFields = valueFieldNames.map((name) => byName.get(name));
  const valueFieldSettings = new Map(
    (opts.valueFieldSettings ?? []).map((settings) => [settings.fieldName, settings]),
  );
  const filterItemFields = new Set((opts.filterItems ?? []).map((item) => item.fieldName));
  const emptyReport =
    !opts.rowField &&
    valueFieldNames.length === 0 &&
    !opts.columnField &&
    filterFieldNames.length === 0;
  // Rows and columns with no values: the layout Recommended PivotTables inserts.
  const layoutOnly =
    rowField !== undefined && columnField !== undefined && valueFieldNames.length === 0;
  if (
    (!emptyReport && !layoutOnly && (!rowField || valueFields.length === 0)) ||
    (opts.rowField && !rowField) ||
    valueFields.some((field) => field === undefined) ||
    (opts.columnField && !columnField) ||
    filterFields.some((field) => field === undefined) ||
    Array.from(filterItemFields).some((name) => !filterFieldNames.includes(name)) ||
    Array.from(pivotFilterFields).some(
      (name) => name !== rowField?.name && name !== columnField?.name,
    ) ||
    (rowField !== undefined && rowField.name === columnField?.name) ||
    Array.from(valueFieldSettings.values()).some(
      (settings) =>
        settings.showAsBaseField !== undefined &&
        settings.showAsBaseField !== rowField?.name &&
        settings.showAsBaseField !== columnField?.name,
    ) ||
    filterFields.some((field) => field?.name === rowField?.name) ||
    (columnField !== undefined && filterFields.some((field) => field?.name === columnField.name))
  ) {
    return refuse('invalid-field');
  }

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
  const roles = new Map<string, PivotAxis>();
  if (rowField) roles.set(rowField.name, PivotAxis.Row);
  if (columnField) roles.set(columnField.name, PivotAxis.Col);
  for (const filterField of filterFields) {
    if (filterField) roles.set(filterField.name, PivotAxis.Page);
  }
  for (const valueField of valueFields) {
    if (valueField && !roles.has(valueField.name)) roles.set(valueField.name, PivotAxis.Value);
  }
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
    if (!filterField) return failAfterPivot('filter-field');
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
    if (!valueField) return failAfterPivot('value-field');
    const settings = valueFieldSettings.get(valueField.name);
    const aggregation =
      settings?.aggregation ??
      opts.aggregation ??
      (valueField.numeric ? PivotAggregation.Sum : PivotAggregation.Count);
    const numberFormat =
      settings?.numberFormat?.trim() ||
      opts.valueNumberFormat?.trim() ||
      (settings?.showValuesAs !== undefined && PERCENT_SHOW_AS.has(settings.showValuesAs)
        ? PERCENT_FORMAT
        : undefined);

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

export function refreshPivotCacheFromRange(
  wb: WorkbookHandle,
  opts: RefreshPivotCacheOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };
  if (!writePivotCacheWorksheetSource(wb, opts.cacheId, opts.source)) {
    return { ok: false, reason: 'engine-failed', step: 'cache-source' };
  }
  const refused = wb.refreshPivotCache(opts.cacheId);
  return refused === null ? { ok: true, cacheId: opts.cacheId } : refreshFailure(refused);
}

export function refreshPivotTableFromRange(
  wb: WorkbookHandle,
  opts: RefreshPivotTableOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };
  const cacheId = wb.pivotTableCacheId(opts.sheet, opts.pivotIndex);
  if (cacheId < 0) return { ok: false, reason: 'invalid-pivot' };
  return refreshPivotCacheFromRange(wb, { cacheId, source: opts.source });
}

export function refreshPivotTable(
  wb: WorkbookHandle,
  opts: RefreshPivotTableFromStoredSourceOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };
  const cacheId = wb.pivotTableCacheId(opts.sheet, opts.pivotIndex);
  if (cacheId < 0) return { ok: false, reason: 'invalid-pivot' };
  const source = pivotCacheWorksheetSourceRange(wb, cacheId);
  if (!source) return { ok: false, reason: 'invalid-range' };
  return refreshPivotCacheFromRange(wb, { cacheId, source });
}

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

/** Rebuild a report on its original pivot cache from a stored definition.
 *  Pivot field `i` binds to cache field `i`, so every field is re-added in
 *  order; a field that sat on no axis is taken back off its axis. The engine
 *  cannot express a field with neither the automatic subtotal nor custom ones,
 *  so that state returns as automatic. Returns the new pivot index, or -1 on
 *  failure. */
export function recreatePivotTable(
  wb: WorkbookHandle,
  sheet: number,
  anchor: { row: number; col: number },
  definition: PivotDefinition,
): number {
  const pivotIndex = wb.createPivotTable(sheet, definition.name, definition.cacheId, anchor);
  if (pivotIndex < 0) return -1;
  const rebuild = (): boolean => {
    const { grandTotals, layout } = definition;
    if (!wb.setPivotTableGrandTotals(sheet, pivotIndex, grandTotals.rows, grandTotals.cols)) {
      return false;
    }
    if (layout !== null && !wb.setPivotReportLayout(sheet, pivotIndex, layout)) return false;
    for (const field of definition.fields) {
      const added = wb.addPivotField(sheet, pivotIndex, {
        sourceName: field.sourceName,
        ...(field.customName ? { customName: field.customName } : {}),
        axis: field.axis ?? PivotAxis.Value,
        subtotalTop: field.subtotalTop,
        ...(field.numberFormat ? { numberFormat: field.numberFormat } : {}),
      });
      if (added !== field.index) return false;
      if (field.axis === null && !wb.setPivotFieldAxis(sheet, pivotIndex, added, null)) {
        return false;
      }
      for (const fn of field.subtotalFns) {
        if (!wb.addPivotFieldSubtotalFn(sheet, pivotIndex, added, fn)) return false;
      }
      if (!wb.setPivotFieldSort(sheet, pivotIndex, added, field.sortOrder, field.sortByField)) {
        return false;
      }
      // A group field comes back with the grouping of the field it groups.
      if (field.groupBaseIndex === null && field.dateGroup) {
        const { granularity, calendar, ...bounds } = field.dateGroup;
        if (!wb.setPivotFieldDateGroup(sheet, pivotIndex, added, granularity, calendar, bounds)) {
          return false;
        }
      }
      if (
        field.groupBaseIndex === null &&
        field.numericGroup &&
        !wb.setPivotFieldNumericGroup(sheet, pivotIndex, added, field.numericGroup)
      ) {
        return false;
      }
      if (field.groupBaseIndex !== null) {
        // A group field's items come back with the grouping; only their visibility is the report's.
        const hidden = field.items.flatMap((item, i) => (item.visible ? [] : [i]));
        if (!hidden.every((i) => wb.setPivotFieldItemVisible(sheet, pivotIndex, added, i, false))) {
          return false;
        }
        continue;
      }
      // A field lists its items when it joins an axis; the saved list replaces them, order included.
      if (!wb.clearPivotFieldItems(sheet, pivotIndex, added)) return false;
      for (const item of field.items) {
        const placed =
          item.cacheIndex >= 0 && wb.capabilities.pivotItemByCacheIndex
            ? wb.addPivotFieldItemAt(sheet, pivotIndex, added, item.cacheIndex, item.visible)
            : wb.addPivotFieldItem(sheet, pivotIndex, added, item.label, item.visible);
        if (!placed) return false;
      }
    }
    // The data fields go in first: a field order places the Values field only once there are two.
    return (
      definition.dataFields.every((spec) => wb.addPivotDataField(sheet, pivotIndex, spec) >= 0) &&
      wb.setPivotRowFieldOrder(sheet, pivotIndex, definition.rowFieldOrder) &&
      wb.setPivotColFieldOrder(sheet, pivotIndex, definition.colFieldOrder) &&
      definition.filters.every((filter) => wb.addPivotFilter(sheet, pivotIndex, filter)) &&
      (definition.style === null ||
        wb.setPivotStyleInfo(sheet, pivotIndex, {
          ...definition.style,
          name: definition.style.name || undefined,
        }))
    );
  };
  if (rebuild()) return pivotIndex;
  wb.removePivotTable(sheet, pivotIndex);
  return -1;
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
