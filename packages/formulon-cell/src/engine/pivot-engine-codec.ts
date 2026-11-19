import type { ErrorCode, NativeValue, PivotFilter, PivotFilterInput } from '@libraz/formulon';
import { PivotDatePeriod } from '@libraz/formulon';
import type { CellValue, PivotAxis, PivotDataFieldSpec, PivotFilterSpec } from './types.js';
import { PivotDateGrouping, ValueKind } from './types.js';
import { formatError } from './value.js';

export function fromNativeValue(v: NativeValue): CellValue {
  if (v === null) return { kind: 'blank' };
  if (typeof v === 'number') return { kind: 'number', value: v };
  if (typeof v === 'boolean') return { kind: 'bool', value: v };
  if (typeof v === 'string') return { kind: 'text', value: v };
  if ('code' in v) return { kind: 'error', code: v.code, text: formatError(v.code, '#ERR!') };
  return { kind: 'blank' };
}

export function pivotCellItemLabel(value: CellValue): string {
  if (value.kind === 'text') return value.value.trim();
  if (value.kind === 'number') return String(value.value);
  if (value.kind === 'bool') return value.value ? 'TRUE' : 'FALSE';
  if (value.kind === 'error') return `#${value.code}`;
  return '';
}

/** A cell-layer value as a pivot cache value. */
export function toCacheValue(value: CellValue):
  | number
  | string
  | boolean
  | null
  | {
      kind: typeof ValueKind.Error;
      errorCode: ErrorCode;
    } {
  if (value.kind === 'blank') return null;
  if (value.kind === 'error') {
    return { kind: ValueKind.Error, errorCode: value.code as ErrorCode };
  }
  return value.value;
}

/** The period bit each cell-layer granularity groups by. */
export const DATE_GROUPING_PERIOD: Record<PivotDateGrouping, number> = {
  [PivotDateGrouping.Day]: PivotDatePeriod.Days,
  [PivotDateGrouping.Month]: PivotDatePeriod.Months,
  [PivotDateGrouping.Quarter]: PivotDatePeriod.Quarters,
  [PivotDateGrouping.Year]: PivotDatePeriod.Years,
  [PivotDateGrouping.Days]: PivotDatePeriod.Days,
  [PivotDateGrouping.Hour]: PivotDatePeriod.Hours,
  [PivotDateGrouping.Minute]: PivotDatePeriod.Minutes,
  [PivotDateGrouping.Second]: PivotDatePeriod.Seconds,
};

/** The granularity of a date group's finest period; null for an ungrouped field. */
export function dateGroupingOf(periods: number): PivotDateGrouping | null {
  if (periods & PivotDatePeriod.Seconds) return PivotDateGrouping.Second;
  if (periods & PivotDatePeriod.Minutes) return PivotDateGrouping.Minute;
  if (periods & PivotDatePeriod.Hours) return PivotDateGrouping.Hour;
  if (periods & PivotDatePeriod.Days) return PivotDateGrouping.Day;
  if (periods & PivotDatePeriod.Months) return PivotDateGrouping.Month;
  if (periods & PivotDatePeriod.Quarters) return PivotDateGrouping.Quarter;
  if (periods & PivotDatePeriod.Years) return PivotDateGrouping.Year;
  return null;
}

export function toDataFieldInput(spec: PivotDataFieldSpec) {
  return {
    // An absent name lets the engine name the field; "" is an empty caption, as the spreadsheet saves it.
    ...(spec.name !== undefined ? { name: spec.name } : {}),
    fieldIndex: spec.fieldIndex,
    aggregation: spec.aggregation,
    ...(spec.numberFormat !== undefined ? { numberFormat: spec.numberFormat.trim() || null } : {}),
    ...(spec.showValuesAs !== undefined ? { showAs: spec.showValuesAs } : {}),
    ...(spec.showAsBaseField !== undefined ? { showAsBaseField: spec.showAsBaseField } : {}),
    ...(spec.showAsBaseItem !== undefined ? { showAsBaseItem: spec.showAsBaseItem } : {}),
  };
}

/** A cell-layer filter spec as the engine's filter input; the field is named
 *  by its source name and the axis is the field's own. */
export function toFilterInput(spec: PivotFilterSpec): PivotFilterInput {
  return {
    fieldIndex: spec.fieldName,
    kind: spec.kind,
    ...(spec.dataFieldIndex !== undefined ? { dataFieldIndex: spec.dataFieldIndex } : {}),
    ...(spec.value !== undefined ? { value: spec.value } : {}),
    ...(spec.valueHigh !== undefined ? { valueHigh: spec.valueHigh } : {}),
  };
}

/** A stored filter in the cell layer's shape; null when its field has no axis or name. */
export function toFilterSpec(
  filter: PivotFilter,
  axis: PivotAxis | null,
  fieldName: string,
): PivotFilterSpec | null {
  if (axis === null || fieldName.trim().length === 0) return null;
  return {
    axis,
    fieldName: fieldName.trim(),
    kind: filter.kind,
    ...(filter.dataFieldIndex > 0 ? { dataFieldIndex: filter.dataFieldIndex } : {}),
    ...(filter.value !== null ? { value: filter.value } : {}),
    ...(filter.valueHigh !== null ? { valueHigh: filter.valueHigh } : {}),
  };
}
