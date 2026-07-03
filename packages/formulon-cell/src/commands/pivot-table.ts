import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { CellValue, PivotFilterSpec, PivotShowValuesAs, Range } from '../engine/types.js';
import { PivotAggregation, PivotAxis } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { addSheet } from './sheet-mutate.js';

export interface PivotSourceField {
  name: string;
  index: number;
  numericCount: number;
}

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
}

export interface CreatePivotTableOptions {
  source: Range;
  destination: { sheet: number; row: number; col: number };
  name?: string;
  rowField: string;
  columnField?: string;
  filterField?: string;
  filterFields?: readonly string[];
  filterItems?: readonly PivotFieldItemVisibility[];
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
}

export type CreatePivotTableResult =
  | { ok: true; cacheId: number; pivotIndex: number }
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

const uniqueFieldName = (raw: string, seen: Map<string, number>): string => {
  const base = raw.trim() || 'Column';
  const n = seen.get(base) ?? 0;
  seen.set(base, n + 1);
  return n === 0 ? base : `${base}${n + 1}`;
};

const valueKey = (value: CellValue): string => {
  if (value.kind === 'number') return `n:${value.value}`;
  if (value.kind === 'text') return `t:${value.value}`;
  if (value.kind === 'bool') return `b:${value.value}`;
  if (value.kind === 'error') return `e:${value.code}`;
  return 'blank';
};

const pivotCacheRecordValue = (
  value: CellValue,
  sharedItemIndexes: Map<string, number>,
): CellValue => {
  if (value.kind !== 'number') return value;
  const index = sharedItemIndexes.get(valueKey(value));
  return index === undefined ? value : { kind: 'number', value: index };
};

const pivotAggregationName = (aggregation: PivotAggregation): string => {
  if (aggregation === PivotAggregation.Count) return 'Count';
  if (aggregation === PivotAggregation.Average) return 'Average';
  if (aggregation === PivotAggregation.Max) return 'Max';
  if (aggregation === PivotAggregation.Min) return 'Min';
  return 'Sum';
};

const MAX_PIVOT_SOURCE_CELLS = 100_000;

const rangeArea = (range: Range): number => (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

const canMaterializePivotSource = (range: Range): boolean =>
  rangeArea(range) <= MAX_PIVOT_SOURCE_CELLS;

const columnName = (col: number): string => {
  let n = col + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

const rangeRef = (range: Range): string =>
  `${columnName(range.c0)}${range.r0 + 1}:${columnName(range.c1)}${range.r1 + 1}`;

const writePivotCacheWorksheetSource = (
  wb: WorkbookHandle,
  cacheId: number,
  source: Range,
): boolean => {
  if (!wb.capabilities.pivotCacheSource || typeof wb.setPivotCacheWorksheetSource !== 'function') {
    return true;
  }
  return wb.setPivotCacheWorksheetSource(cacheId, {
    present: true,
    ref: rangeRef(source),
    sheet: wb.sheetName(source.sheet),
  });
};

const sheetIndexByName = (wb: WorkbookHandle, name: string): number => {
  const target = name.trim().toLowerCase();
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (wb.sheetName(sheet).toLowerCase() === target) return sheet;
  }
  return -1;
};

const pivotCacheWorksheetSourceRange = (wb: WorkbookHandle, cacheId: number): Range | null => {
  const source = wb.getPivotCacheWorksheetSource(cacheId);
  if (!source?.present || !source.ref) return null;
  const parsed = parseRangeRef(
    source.sheet ? `${quoteSheetName(source.sheet)}!${source.ref}` : source.ref,
  );
  if (!parsed) return null;
  const sheet =
    parsed.sheetName !== null
      ? sheetIndexByName(wb, parsed.sheetName)
      : source.sheet
        ? sheetIndexByName(wb, source.sheet)
        : -1;
  if (sheet < 0) return null;
  return { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 };
};

const quoteSheetName = (name: string): string =>
  /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

export function inferPivotSourceFields(wb: WorkbookHandle, range: Range): PivotSourceField[] {
  if (range.r1 <= range.r0 || range.c1 < range.c0) return [];
  if (!canMaterializePivotSource(range)) return [];
  const seen = new Map<string, number>();
  const fields: PivotSourceField[] = [];
  for (let c = range.c0; c <= range.c1; c += 1) {
    const fallback = `Column${c - range.c0 + 1}`;
    const name = uniqueFieldName(
      cellLabel(wb.getValue({ sheet: range.sheet, row: range.r0, col: c }), fallback),
      seen,
    );
    let numericCount = 0;
    for (let r = range.r0 + 1; r <= range.r1; r += 1) {
      if (wb.getValue({ sheet: range.sheet, row: r, col: c }).kind === 'number') numericCount += 1;
    }
    fields.push({ name, index: c - range.c0, numericCount });
  }
  return fields;
}

export function inferPivotFieldItems(
  wb: WorkbookHandle,
  range: Range,
  fieldName: string,
): string[] {
  const field = inferPivotSourceFields(wb, range).find((candidate) => candidate.name === fieldName);
  if (!field) return [];
  const seen = new Set<string>();
  const items: string[] = [];
  for (let r = range.r0 + 1; r <= range.r1; r += 1) {
    const value = wb.getValue({
      sheet: range.sheet,
      row: r,
      col: range.c0 + field.index,
    });
    const label = cellLabel(value, '(blank)');
    if (seen.has(label)) continue;
    seen.add(label);
    items.push(label);
  }
  return items;
}

export function createPivotTableFromRange(
  wb: WorkbookHandle,
  opts: CreatePivotTableOptions,
): CreatePivotTableResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };

  const fields = inferPivotSourceFields(wb, opts.source);
  if (fields.length < 2) return { ok: false, reason: 'invalid-range' };
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
  if (
    !rowField ||
    valueFields.length === 0 ||
    valueFields.some((field) => field === undefined) ||
    (opts.columnField && !columnField) ||
    filterFields.some((field) => field === undefined) ||
    Array.from(filterItemFields).some((name) => !filterFieldNames.includes(name)) ||
    Array.from(pivotFilterFields).some((name) => !filterFieldNames.includes(name)) ||
    rowField.name === columnField?.name ||
    filterFields.some((field) => field?.name === rowField.name) ||
    (columnField !== undefined && filterFields.some((field) => field?.name === columnField.name))
  ) {
    return { ok: false, reason: 'invalid-field' };
  }

  const cacheId = wb.createPivotCache(0);
  if (cacheId < 0) return { ok: false, reason: 'engine-failed', step: 'cache' };
  if (!writePivotCacheWorksheetSource(wb, cacheId, opts.source)) {
    wb.removePivotCache(cacheId);
    return { ok: false, reason: 'engine-failed', step: 'cache-source' };
  }

  for (const field of fields) {
    if (wb.addPivotCacheField(cacheId, field.name) < 0) {
      wb.removePivotCache(cacheId);
      return { ok: false, reason: 'engine-failed', step: 'cache-field' };
    }
  }

  const sharedItemIndexesByField = new Map<number, Map<string, number>>();
  for (const field of fields) {
    const seen = new Set<string>();
    const indexes = new Map<string, number>();
    for (let r = opts.source.r0 + 1; r <= opts.source.r1; r += 1) {
      const value = wb.getValue({
        sheet: opts.source.sheet,
        row: r,
        col: opts.source.c0 + field.index,
      });
      const key = valueKey(value);
      if (seen.has(key)) continue;
      seen.add(key);
      indexes.set(key, indexes.size);
      if (!wb.addPivotCacheSharedItem(cacheId, field.index, value)) {
        wb.removePivotCache(cacheId);
        return { ok: false, reason: 'engine-failed', step: 'shared-item' };
      }
    }
    sharedItemIndexesByField.set(field.index, indexes);
  }

  for (let r = opts.source.r0 + 1; r <= opts.source.r1; r += 1) {
    const recordIdx = wb.addPivotCacheRecord(cacheId);
    if (recordIdx < 0) {
      wb.removePivotCache(cacheId);
      return { ok: false, reason: 'engine-failed', step: 'cache-record' };
    }
    for (let c = opts.source.c0; c <= opts.source.c1; c += 1) {
      const ok = wb.setPivotCacheRecordValue(
        cacheId,
        recordIdx,
        c - opts.source.c0,
        pivotCacheRecordValue(
          wb.getValue({ sheet: opts.source.sheet, row: r, col: c }),
          sharedItemIndexesByField.get(c - opts.source.c0) ?? new Map(),
        ),
      );
      if (!ok) {
        wb.removePivotCache(cacheId);
        return { ok: false, reason: 'engine-failed', step: 'cache-value' };
      }
    }
  }

  const pivotIndex = wb.createPivotTable(
    opts.destination.sheet,
    opts.name?.trim() || 'PivotTable1',
    cacheId,
    { row: opts.destination.row, col: opts.destination.col },
  );
  if (pivotIndex < 0) {
    wb.removePivotCache(cacheId);
    return { ok: false, reason: 'engine-failed', step: 'pivot' };
  }
  const failAfterPivot = (step: string): CreatePivotTableResult => {
    wb.removePivotTable(opts.destination.sheet, pivotIndex);
    wb.removePivotCache(cacheId);
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

  const rowPivotField = wb.addPivotField(opts.destination.sheet, pivotIndex, {
    sourceName: rowField.name,
    axis: PivotAxis.Row,
    subtotalTop: opts.rowSubtotalTop ?? true,
  });
  if (rowPivotField < 0) return failAfterPivot('row-field');
  if (opts.rowSort && opts.rowSort !== 'none') {
    const ok = wb.setPivotFieldSort(
      opts.destination.sheet,
      pivotIndex,
      rowPivotField,
      opts.rowSort === 'asc',
      '',
    );
    if (!ok) return failAfterPivot('row-sort');
  }

  if (columnField) {
    const colPivotField = wb.addPivotField(opts.destination.sheet, pivotIndex, {
      sourceName: columnField.name,
      axis: PivotAxis.Col,
      subtotalTop: opts.columnSubtotalTop ?? true,
    });
    if (colPivotField < 0) return failAfterPivot('col-field');
    if (opts.columnSort && opts.columnSort !== 'none') {
      const ok = wb.setPivotFieldSort(
        opts.destination.sheet,
        pivotIndex,
        colPivotField,
        opts.columnSort === 'asc',
        '',
      );
      if (!ok) return failAfterPivot('col-sort');
    }
  }

  for (const filterField of filterFields) {
    if (!filterField) return failAfterPivot('filter-field');
    const filterPivotField = wb.addPivotField(opts.destination.sheet, pivotIndex, {
      sourceName: filterField.name,
      axis: PivotAxis.Page,
    });
    if (filterPivotField < 0) return failAfterPivot('filter-field');
    for (const item of opts.filterItems?.filter((entry) => entry.fieldName === filterField.name) ??
      []) {
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
    for (const filter of opts.pivotFilters?.filter(
      (entry) => entry.fieldName === filterField.name,
    ) ?? []) {
      if (!wb.addPivotFilter(opts.destination.sheet, pivotIndex, filter)) {
        return failAfterPivot('pivot-filter');
      }
    }
  }

  for (const valueField of valueFields) {
    if (!valueField) return failAfterPivot('value-field');
    const valuePivotField = wb.addPivotField(opts.destination.sheet, pivotIndex, {
      sourceName: valueField.name,
      axis: PivotAxis.Value,
    });
    if (valuePivotField < 0) return failAfterPivot('value-field');
    const settings = valueFieldSettings.get(valueField.name);
    const aggregation =
      settings?.aggregation ??
      opts.aggregation ??
      (valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count);
    const numberFormat = settings?.numberFormat?.trim() || opts.valueNumberFormat?.trim();

    const dataField = wb.addPivotDataField(opts.destination.sheet, pivotIndex, {
      name: `${pivotAggregationName(aggregation)} of ${valueField.name}`,
      fieldIndex: valueField.index,
      aggregation,
      numberFormat: numberFormat || undefined,
      showValuesAs: settings?.showValuesAs,
    });
    if (dataField < 0) return failAfterPivot('data-field');
  }

  return { ok: true, cacheId, pivotIndex };
}

export function refreshPivotCacheFromRange(
  wb: WorkbookHandle,
  opts: RefreshPivotCacheOptions,
): RefreshPivotCacheResult {
  if (!wb.capabilities.pivotTableMutate) return { ok: false, reason: 'unsupported' };

  const fields = inferPivotSourceFields(wb, opts.source);
  if (fields.length < 2) return { ok: false, reason: 'invalid-range' };
  const cacheFields = wb.pivotCacheFieldNames(opts.cacheId);
  if (
    cacheFields.length !== fields.length ||
    fields.some((field, index) => field.name !== cacheFields[index])
  ) {
    return { ok: false, reason: 'invalid-field' };
  }

  const sharedItemIndexesByField = new Map<number, Map<string, number>>();
  for (const field of fields) {
    if (!wb.clearPivotCacheSharedItems(opts.cacheId, field.index)) {
      return { ok: false, reason: 'engine-failed', step: 'shared-item-clear' };
    }
    const seen = new Set<string>();
    const indexes = new Map<string, number>();
    for (let r = opts.source.r0 + 1; r <= opts.source.r1; r += 1) {
      const value = wb.getValue({
        sheet: opts.source.sheet,
        row: r,
        col: opts.source.c0 + field.index,
      });
      const key = valueKey(value);
      if (seen.has(key)) continue;
      seen.add(key);
      indexes.set(key, indexes.size);
      if (!wb.addPivotCacheSharedItem(opts.cacheId, field.index, value)) {
        return { ok: false, reason: 'engine-failed', step: 'shared-item' };
      }
    }
    sharedItemIndexesByField.set(field.index, indexes);
  }

  if (!wb.clearPivotCacheRecords(opts.cacheId)) {
    return { ok: false, reason: 'engine-failed', step: 'cache-record-clear' };
  }
  for (let r = opts.source.r0 + 1; r <= opts.source.r1; r += 1) {
    const recordIdx = wb.addPivotCacheRecord(opts.cacheId);
    if (recordIdx < 0) return { ok: false, reason: 'engine-failed', step: 'cache-record' };
    for (let c = opts.source.c0; c <= opts.source.c1; c += 1) {
      const ok = wb.setPivotCacheRecordValue(
        opts.cacheId,
        recordIdx,
        c - opts.source.c0,
        pivotCacheRecordValue(
          wb.getValue({ sheet: opts.source.sheet, row: r, col: c }),
          sharedItemIndexesByField.get(c - opts.source.c0) ?? new Map(),
        ),
      );
      if (!ok) return { ok: false, reason: 'engine-failed', step: 'cache-record-value' };
    }
  }

  if (!writePivotCacheWorksheetSource(wb, opts.cacheId, opts.source)) {
    return { ok: false, reason: 'engine-failed', step: 'cache-source' };
  }

  return { ok: true, cacheId: opts.cacheId };
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

export type RibbonPivotTableAction =
  | 'dialog'
  | 'recommended'
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
  action: RibbonPivotTableAction;
  strings: RibbonPivotTableActionStrings;
  history?: History | null;
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
    if (!pivot) {
      return {
        kind: 'report',
        report: {
          title: strings.pivotTableRefreshData,
          items: [
            {
              severity: 'warning',
              label: strings.pivotTable,
              detail: strings.pivotTableRefreshUnavailable,
            },
          ],
        },
      };
    }
    const result = refreshPivotTable(workbook, {
      sheet: pivot.sheetIndex,
      pivotIndex: pivot.pivotIndex,
    });
    if (!result.ok) {
      return {
        kind: 'report',
        report: {
          title: strings.pivotTableRefreshData,
          items: [
            {
              severity: 'warning',
              label: strings.pivotTable,
              detail: strings.pivotTableRefreshUnavailable,
            },
          ],
        },
      };
    }
    mutators.replaceCells(store, workbook.cells(pivot.sheetIndex));
    return { kind: 'refreshed', sheet: pivot.sheetIndex };
  }
  if (action === 'dialog' || action === 'existing-sheet') {
    return { kind: 'open-dialog' };
  }
  if (action === 'recommended') {
    return {
      kind: 'report',
      report: {
        title: strings.recommendedPivotTables,
        items: [
          { severity: 'info', label: strings.pivotTable, detail: strings.pivotAuthoringDetail },
        ],
      },
    };
  }
  if (!workbook.capabilities.pivotTableMutate) {
    return {
      kind: 'report',
      report: {
        title: strings.pivotTableNewSheet,
        items: [
          { severity: 'info', label: strings.pivotTable, detail: strings.pivotAuthoringDetail },
        ],
      },
    };
  }
  const source = store.getState().selection.range;
  const fields = inferPivotSourceFields(workbook, source);
  const valueField = fields.find((field) => field.numericCount > 0) ?? fields.at(-1);
  const rowField = fields.find((field) => field.name !== valueField?.name) ?? fields[0];
  if (!rowField || !valueField || rowField.name === valueField.name) {
    return {
      kind: 'report',
      report: {
        title: strings.pivotTable,
        items: [
          { severity: 'warning', label: strings.pivotTable, detail: strings.pivotAuthoringDetail },
        ],
      },
    };
  }
  let destinationSheet = source.sheet;
  if (action === 'new-sheet') {
    const added = addSheet(store, workbook, history);
    if (added < 0) {
      return {
        kind: 'report',
        report: {
          title: strings.pivotTableNewSheet,
          items: [
            {
              severity: 'warning',
              label: strings.pivotTable,
              detail: strings.workbookStructureProtectedBlocked,
            },
          ],
        },
      };
    }
    destinationSheet = added;
  }
  const destination =
    action === 'new-sheet'
      ? { sheet: destinationSheet, row: 0, col: 0 }
      : { sheet: destinationSheet, row: source.r1 + 3, col: source.c0 };
  const result = createPivotTableFromRange(workbook, {
    source,
    destination,
    name: `PivotTable${workbook.getPivotTables().length + 1}`,
    rowField: rowField.name,
    valueField: valueField.name,
    aggregation: valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count,
  });
  if (result.ok) {
    mutators.replaceCells(store, workbook.cells(destinationSheet));
    mutators.setSheetIndex(store, destinationSheet);
    mutators.setActive(store, destination);
    return { kind: 'created', destinationSheet, destination };
  }
  return {
    kind: 'report',
    report: {
      title: strings.pivotTableNewSheet,
      items: [
        { severity: 'info', label: strings.pivotTable, detail: strings.pivotAuthoringDetail },
      ],
    },
  };
};
