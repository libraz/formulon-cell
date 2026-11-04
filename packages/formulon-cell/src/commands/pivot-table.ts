import { formatA1Range } from '../engine/address.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import { pivotStyleLayer } from '../engine/pivot-style.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { CellValue, PivotFilterSpec, Range } from '../engine/types.js';
import {
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotAxis,
  PivotShowValuesAs,
  PivotSortOrder,
} from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { PivotDefinition } from '../engine/workbook-handle-pivot.js';
import { rangeArea } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { AutofitOptions } from './autofit-measurement.js';
import { currentRegion } from './current-region.js';
import { defaultObjectNames } from './format-as-table.js';
import type { History } from './history.js';
import { reflowPivotFormats } from './pivot-format-reflow.js';
import { autofitColsWidthForRanges } from './row-col-layout.js';
import { addSheet, moveSheet, removeSheet } from './sheet-mutate.js';

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
  /** Position in `inferPivotFieldItems` of the base field, or a base-item sentinel. */
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

const MAX_PIVOT_SOURCE_CELLS = 100_000;

const canMaterializePivotSource = (range: Range): boolean =>
  rangeArea(range) <= MAX_PIVOT_SOURCE_CELLS;

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
    ref: formatA1Range(source, { collapse: false }),
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

/** Whether the worksheet column behind a cache field holds numbers. False
 *  when the cache has no readable worksheet source. */
export function pivotCacheFieldIsNumeric(
  wb: WorkbookHandle,
  cacheId: number,
  cacheFieldIndex: number,
): boolean {
  const range = pivotCacheWorksheetSourceRange(wb, cacheId);
  if (!range) return false;
  const field = inferPivotSourceFields(wb, range)[cacheFieldIndex];
  return (field?.numericCount ?? 0) > 0;
}

export function inferPivotFieldItems(
  wb: WorkbookHandle,
  range: Range,
  fieldName: string,
  fields: readonly PivotSourceField[] = inferPivotSourceFields(wb, range),
): string[] {
  const field = fields.find((candidate) => candidate.name === fieldName);
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
  const emptyReport =
    !opts.rowField &&
    valueFieldNames.length === 0 &&
    !opts.columnField &&
    filterFieldNames.length === 0;
  if (
    (!emptyReport && (!rowField || valueFields.length === 0)) ||
    (opts.rowField && !rowField) ||
    valueFields.some((field) => field === undefined) ||
    (opts.columnField && !columnField) ||
    filterFields.some((field) => field === undefined) ||
    Array.from(filterItemFields).some((name) => !filterFieldNames.includes(name)) ||
    Array.from(pivotFilterFields).some((name) => !filterFieldNames.includes(name)) ||
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
    opts.name?.trim() || nextPivotTableName(wb),
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

  for (const filterField of filterFields) {
    if (!filterField) return failAfterPivot('filter-field');
    const filterPivotField = pivotFieldByName.get(filterField.name) as number;
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

  // A specific base item is a position in the base field's items, so list them first.
  const baseFieldIndexes = new Map<string, number>();
  if (rowField) baseFieldIndexes.set(rowField.name, rowPivotField);
  if (columnField) baseFieldIndexes.set(columnField.name, colPivotField);
  const itemsListed = new Set<string>();
  for (const settings of valueFieldSettings.values()) {
    const baseName = settings.showAsBaseField;
    const baseItem = settings.showAsBaseItem;
    if (
      baseName === undefined ||
      baseItem === undefined ||
      baseItem === PIVOT_SHOW_AS_BASE_PREVIOUS ||
      baseItem === PIVOT_SHOW_AS_BASE_NEXT ||
      itemsListed.has(baseName)
    ) {
      continue;
    }
    itemsListed.add(baseName);
    for (const item of inferPivotFieldItems(wb, opts.source, baseName)) {
      if (
        !wb.addPivotFieldItem(
          opts.destination.sheet,
          pivotIndex,
          baseFieldIndexes.get(baseName) as number,
          item,
          true,
        )
      ) {
        return failAfterPivot('base-item');
      }
    }
  }

  for (const valueField of valueFields) {
    if (!valueField) return failAfterPivot('value-field');
    const settings = valueFieldSettings.get(valueField.name);
    const aggregation =
      settings?.aggregation ??
      opts.aggregation ??
      (valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count);
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

  return { ok: true, cacheId, pivotIndex };
}

/** Labels of each cache field's items in cache-index order, read from a
 *  report drawing on the cache (the engine cannot list shared items). */
const cachedItemLabels = (wb: WorkbookHandle, cacheId: number): Map<number, string[]> => {
  const out = new Map<number, string[]>();
  const report = wb
    .getPivotTables()
    .find((p) => wb.pivotTableCacheId(p.sheetIndex, p.pivotIndex) === cacheId);
  if (!report) return out;
  const { sheetIndex: sheet, pivotIndex } = report;
  const count = wb.pivotFieldCount(sheet, pivotIndex);
  for (let fieldIdx = 0; fieldIdx < count; fieldIdx += 1) {
    const items = wb.getPivotField(sheet, pivotIndex, fieldIdx)?.items ?? [];
    const labels: string[] = [];
    for (const item of items) if (item.cacheIndex >= 0) labels[item.cacheIndex] = item.label;
    if (labels.length > 0) out.set(fieldIdx, labels);
  }
  return out;
};

/** Position of `value` among previously cached labels; unseen values sort after. */
const rankByLabel = (labels: readonly string[] | undefined, value: CellValue): number => {
  const label =
    value.kind === 'text' ? value.value : value.kind === 'number' ? String(value.value) : null;
  const at = label === null || !labels ? -1 : labels.indexOf(label);
  return at < 0 ? Number.MAX_SAFE_INTEGER : at;
};

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

  const previousOrder = cachedItemLabels(wb, opts.cacheId);
  const sharedItemIndexesByField = new Map<number, Map<string, number>>();
  for (const field of fields) {
    if (!wb.clearPivotCacheSharedItems(opts.cacheId, field.index)) {
      return { ok: false, reason: 'engine-failed', step: 'shared-item-clear' };
    }
    const values = new Map<string, CellValue>();
    for (let r = opts.source.r0 + 1; r <= opts.source.r1; r += 1) {
      const value = wb.getValue({
        sheet: opts.source.sheet,
        row: r,
        col: opts.source.c0 + field.index,
      });
      const key = valueKey(value);
      if (!values.has(key)) values.set(key, value);
    }
    // The report's field items point at cache indices, so values the cache
    // held keep their old order and new ones follow, as the desktop app does.
    const ordered = [...values.values()].sort(
      (a, b) =>
        rankByLabel(previousOrder.get(field.index), a) -
        rankByLabel(previousOrder.get(field.index), b),
    );
    const indexes = new Map<string, number>();
    for (const value of ordered) {
      indexes.set(valueKey(value), indexes.size);
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

/** Where a new report lands: a fresh sheet inserted before the source sheet
 *  (report at A3), or a cell of the source sheet. */
/** Fit the columns of the PivotTables on the sheet on show to their content,
 *  as the desktop does after a report update; a column can shrink, and columns
 *  outside the reports keep their width. Part of the update, so no undo step. */
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
    .filter((b) => !empty.has(b.pivotIndex))
    .map((b) => b.range);
  if (ranges.length > 0) autofitColsWidthForRanges(store, null, ranges, wb, opts);
}

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
      reason: 'sheet-failed' | 'unsupported' | 'invalid-range' | 'invalid-field' | 'engine-failed';
      step?: string;
    };

export interface CreatePlacedPivotDeps {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history?: History | null;
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

/** Create a report at `placement` as one undo step (the new sheet, its
 *  position and the report revert together). */
export function createPivotTableAtPlacement(
  deps: CreatePlacedPivotDeps,
  opts: Omit<CreatePivotTableOptions, 'destination'>,
  placement: PivotPlacement,
): CreatePlacedPivotResult {
  const { store, workbook: wb } = deps;
  const history = deps.history && !deps.history.isReplaying() ? deps.history : null;
  const txn = history?.begin();
  let source = opts.source;
  let destination = {
    sheet: source.sheet,
    ...(placement === 'new-sheet' ? PIVOT_NEW_SHEET_ANCHOR : placement),
  };
  let addedSheet = -1;
  if (placement === 'new-sheet') {
    addedSheet = addSheet(store, wb, history);
    if (addedSheet < 0) {
      if (history && txn) history.end(txn);
      return { ok: false, reason: 'sheet-failed' };
    }
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
    if (history && txn) {
      if (addedSheet >= 0) history.abort(txn);
      else history.end(txn);
    } else if (addedSheet >= 0) {
      removeSheet(store, wb, addedSheet);
    }
    return result;
  }
  if (history && txn) {
    let current = result;
    const refresh = (): void => {
      if (store.getState().data.sheetIndex === destination.sheet) {
        mutators.replaceCells(store, wb.cells(destination.sheet));
      }
    };
    history.push({
      undo: () => {
        wb.removePivotTable(destination.sheet, current.pivotIndex);
        wb.removePivotCache(current.cacheId);
        refresh();
      },
      redo: () => {
        const again = create();
        if (again.ok) current = again;
        refresh();
      },
    });
    history.end(txn);
  }
  return { ok: true, cacheId: result.cacheId, pivotIndex: result.pivotIndex, destination };
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
      if (field.dateGroup) {
        const { granularity, calendar, ...bounds } = field.dateGroup;
        if (!wb.setPivotFieldDateGroup(sheet, pivotIndex, added, granularity, calendar, bounds)) {
          return false;
        }
      }
      for (const item of field.items) {
        const placed =
          item.cacheIndex >= 0 && wb.capabilities.pivotItemByCacheIndex
            ? wb.addPivotFieldItemAt(sheet, pivotIndex, added, item.cacheIndex, item.visible)
            : wb.addPivotFieldItem(sheet, pivotIndex, added, item.label, item.visible);
        if (!placed) return false;
      }
    }
    return (
      wb.setPivotRowFieldOrder(sheet, pivotIndex, definition.rowFieldOrder) &&
      wb.setPivotColFieldOrder(sheet, pivotIndex, definition.colFieldOrder) &&
      definition.dataFields.every((spec) => wb.addPivotDataField(sheet, pivotIndex, spec) >= 0) &&
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
  const bounds = wb.pivotBounds(sheet).find((b) => b.pivotIndex === pivotIndex)?.range;
  const definition = bounds ? wb.getPivotDefinition(sheet, pivotIndex) : null;
  if (!bounds || !definition) return;
  const filterRows = definition.fields.filter((f) => f.axis === PivotAxis.Page).length;
  if (filterRows === 0) return;
  const row = bounds.r0 + filterRows + 1;
  wb.setPivotTableAnchor(sheet, pivotIndex, {
    row,
    col: bounds.c0,
    rows: bounds.r1 - row + 1,
    cols: bounds.c1 - bounds.c0 + 1,
  });
}

/** Bring the active sheet's reports up to date after a report edit, an undo
 *  or a redo: settle their anchors, reload the projected cells, move direct
 *  formats with the items they belong to and refit the report columns. */
export function refreshPivotReportsAfterEdit(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  opts?: AutofitOptions,
): void {
  const sheet = store.getState().data.sheetIndex;
  for (const { pivotIndex } of wb.pivotBounds(sheet)) settlePivotAnchor(wb, sheet, pivotIndex);
  mutators.replaceCells(store, wb.cells(sheet));
  reflowPivotFormats(store, wb);
  autofitPivotTableColumns(store, wb, opts);
}

/** The report whose whole extent is exactly `range`, or null. */
export function pivotTableAtRange(wb: WorkbookHandle, range: Range): number | null {
  if (!wb.capabilities.pivotTables) return null;
  const hit = wb
    .pivotBounds(range.sheet)
    .find(
      (b) =>
        b.range.r0 === range.r0 &&
        b.range.c0 === range.c0 &&
        b.range.r1 === range.r1 &&
        b.range.c1 === range.c1,
    );
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
  const { store, workbook: wb } = deps;
  const history = deps.history && !deps.history.isReplaying() ? deps.history : null;
  const definition = wb.getPivotDefinition(source.sheet, source.pivotIndex);
  if (!definition) return null;
  // The extent starts at the filter block, one blank row above the anchor.
  const filterRows = definition.fields.filter((f) => f.axis === PivotAxis.Page).length;
  const anchor = { row: at.row + (filterRows > 0 ? filterRows + 1 : 0), col: at.col };
  const copy = { ...definition, name: nextPivotTableName(wb) };
  let pivotIndex = recreatePivotTable(wb, at.sheet, anchor, copy);
  if (pivotIndex < 0) return null;
  const refresh = (): void => {
    if (store.getState().data.sheetIndex === at.sheet) {
      mutators.replaceCells(store, wb.cells(at.sheet));
    }
  };
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
  return wb.pivotBounds(at.sheet).find((b) => b.pivotIndex === pivotIndex)?.range ?? null;
}

export interface DeletePivotTablesDeps {
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
  const { store, workbook: wb } = deps;
  const history = deps.history && !deps.history.isReplaying() ? deps.history : null;
  const ascending = [...pivots].sort((a, b) => a.pivotIndex - b.pivotIndex);
  const captured: { definition: PivotDefinition; anchor: { row: number; col: number } }[] = [];
  for (const { pivotIndex, range } of ascending) {
    const definition = wb.getPivotDefinition(sheet, pivotIndex);
    if (!definition) return false;
    // The extent starts at the filter block, which sits above the anchor with
    // one blank row between them.
    const filterRows = definition.fields.filter((f) => f.axis === PivotAxis.Page).length;
    captured.push({
      definition,
      anchor: { row: range.r0 + (filterRows > 0 ? filterRows + 1 : 0), col: range.c0 },
    });
  }
  const refresh = (): void => {
    if (store.getState().data.sheetIndex === sheet) {
      mutators.replaceCells(store, wb.cells(sheet));
    }
  };
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
  const source = pivotSourceForSelection(store, workbook);
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
  const result = createPivotTableAtPlacement(
    { store, workbook, history },
    {
      source,
      name: nextPivotTableName(workbook),
      rowField: rowField.name,
      valueField: valueField.name,
      aggregation: valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count,
    },
    'new-sheet',
  );
  if (result.ok) {
    const { destination } = result;
    mutators.replaceCells(store, workbook.cells(destination.sheet));
    mutators.setSheetIndex(store, destination.sheet);
    mutators.setActive(store, destination);
    return { kind: 'created', destinationSheet: destination.sheet, destination };
  }
  if (result.reason === 'sheet-failed') {
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
