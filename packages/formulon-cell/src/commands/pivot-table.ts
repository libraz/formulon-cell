import type { CellValue, Range } from '../engine/types.js';
import { PivotAggregation, PivotAxis } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

export interface PivotSourceField {
  name: string;
  index: number;
  numericCount: number;
}

export type PivotSortDirection = 'none' | 'asc' | 'desc';

export interface CreatePivotTableOptions {
  source: Range;
  destination: { sheet: number; row: number; col: number };
  name?: string;
  rowField: string;
  columnField?: string;
  valueField: string;
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
  const valueField = byName.get(opts.valueField);
  if (
    !rowField ||
    !valueField ||
    (opts.columnField && !columnField) ||
    rowField.name === columnField?.name
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

  return { ok: true, cacheId, pivotIndex };
}
