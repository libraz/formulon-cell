import { PivotAggregation, PivotAxis, PivotShowValuesAs } from '../engine/types.js';
import type { PivotSourceField } from './pivot-cache-source.js';
import type { CreatePivotTableOptions, PivotValueFieldSettings } from './pivot-table.js';

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

/** The cache fields a new report places on each axis, resolved from the creation options. */
export interface PivotFieldPlan {
  rowField: PivotSourceField | undefined;
  columnField: PivotSourceField | undefined;
  filterFields: PivotSourceField[];
  valueFields: PivotSourceField[];
  valueFieldSettings: Map<string, PivotValueFieldSettings>;
}

/** Resolves and validates the report's fields against the cache; a string names the refusal. */
export function resolvePivotFieldPlan(
  fields: PivotSourceField[],
  opts: CreatePivotTableOptions,
): PivotFieldPlan | 'invalid-range' | 'invalid-field' {
  if (fields.length < 2) return 'invalid-range';
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
  const filterFields = filterFieldNames
    .map((name) => byName.get(name))
    .filter((field): field is PivotSourceField => field !== undefined);
  const pivotFilterFields = new Set((opts.pivotFilters ?? []).map((filter) => filter.fieldName));
  const valueFieldNames = Array.from(
    new Set(
      (opts.valueFields && opts.valueFields.length > 0
        ? opts.valueFields
        : [opts.valueField]
      ).filter(Boolean),
    ),
  );
  const valueFields = valueFieldNames
    .map((name) => byName.get(name))
    .filter((field): field is PivotSourceField => field !== undefined);
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
    valueFields.length !== valueFieldNames.length ||
    (opts.columnField && !columnField) ||
    filterFields.length !== filterFieldNames.length ||
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
    filterFields.some((field) => field.name === rowField?.name) ||
    (columnField !== undefined && filterFields.some((field) => field.name === columnField.name))
  ) {
    return 'invalid-field';
  }
  return { rowField, columnField, filterFields, valueFields, valueFieldSettings };
}

/** The axis each used cache field takes; value fields already on an axis keep it. */
export function pivotFieldRoles(plan: PivotFieldPlan): Map<string, PivotAxis> {
  const { rowField, columnField, filterFields, valueFields } = plan;
  const roles = new Map<string, PivotAxis>();
  if (rowField) roles.set(rowField.name, PivotAxis.Row);
  if (columnField) roles.set(columnField.name, PivotAxis.Col);
  for (const filterField of filterFields) {
    roles.set(filterField.name, PivotAxis.Page);
  }
  for (const valueField of valueFields) {
    if (!roles.has(valueField.name)) roles.set(valueField.name, PivotAxis.Value);
  }
  return roles;
}

/** A value field's aggregation and number format; the format may be an empty string. */
export function valueFieldFormat(
  valueField: PivotSourceField,
  settings: PivotValueFieldSettings | undefined,
  opts: CreatePivotTableOptions,
): { aggregation: PivotAggregation; numberFormat: string | undefined } {
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
  return { aggregation, numberFormat };
}
