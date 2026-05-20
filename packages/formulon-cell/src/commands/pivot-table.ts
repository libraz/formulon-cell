import type { CellValue, PivotFilterSpec, Range } from '../engine/types.js';
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

export function inferPivotSourceFields(wb: WorkbookHandle, range: Range): PivotSourceField[] {
  if (range.r1 <= range.r0 || range.c1 < range.c0) return [];
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

  for (const field of fields) {
    if (wb.addPivotCacheField(cacheId, field.name) < 0) {
      wb.removePivotCache(cacheId);
      return { ok: false, reason: 'engine-failed', step: 'cache-field' };
    }
  }

  for (const field of fields) {
    const seen = new Set<string>();
    for (let r = opts.source.r0 + 1; r <= opts.source.r1; r += 1) {
      const value = wb.getValue({
        sheet: opts.source.sheet,
        row: r,
        col: opts.source.c0 + field.index,
      });
      const key = valueKey(value);
      if (seen.has(key)) continue;
      seen.add(key);
      if (!wb.addPivotCacheSharedItem(cacheId, field.index, value)) {
        wb.removePivotCache(cacheId);
        return { ok: false, reason: 'engine-failed', step: 'shared-item' };
      }
    }
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
        wb.getValue({ sheet: opts.source.sheet, row: r, col: c }),
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
    const aggregation =
      opts.aggregation ??
      (valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count);

    const dataField = wb.addPivotDataField(opts.destination.sheet, pivotIndex, {
      name: `${aggregation === PivotAggregation.Count ? 'Count' : 'Sum'} of ${valueField.name}`,
      fieldIndex: valueField.index,
      aggregation,
      numberFormat: opts.valueNumberFormat?.trim() || undefined,
    });
    if (dataField < 0) return failAfterPivot('data-field');
  }

  return { ok: true, cacheId, pivotIndex };
}

export type RibbonPivotTableAction = 'dialog' | 'recommended' | 'new-sheet' | 'existing-sheet';

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
  | { kind: 'report'; report: RibbonPivotTableReport };

export interface RibbonPivotTableActionStrings {
  pivotTable: string;
  pivotTableNewSheet: string;
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
 *    insufficient source data, blocked by structure protection, or engine
 *    refusal to author the pivot).
 *
 *  Encapsulating the branching here keeps React, Vue, and the playground
 *  shells identical and removes ~100 lines of duplication per host. */
export const executeRibbonPivotTableAction = (
  deps: ExecuteRibbonPivotTableActionDeps,
): RibbonPivotTableActionResult => {
  const { store, workbook, action, strings, history = null } = deps;
  if (action === 'dialog' || action === 'existing-sheet') {
    return { kind: 'open-dialog' };
  }
  if (!workbook.capabilities.pivotTableMutate) {
    return {
      kind: 'report',
      report: {
        title:
          action === 'recommended' ? strings.recommendedPivotTables : strings.pivotTableNewSheet,
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
      title: action === 'recommended' ? strings.recommendedPivotTables : strings.pivotTableNewSheet,
      items: [
        { severity: 'info', label: strings.pivotTable, detail: strings.pivotAuthoringDetail },
      ],
    },
  };
};
