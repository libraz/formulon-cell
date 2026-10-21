import { PivotAggregation } from './types.js';

/** Per-aggregation prefix of a value field's default name ("Sum of"). */
export interface PivotDataFieldPrefixes {
  sum: string;
  count: string;
  average: string;
  max: string;
  min: string;
  product: string;
  countNumbers: string;
  stdDev: string;
  stdDevP: string;
  variance: string;
  varianceP: string;
}

/** How a new value field is named in the UI language. */
export interface PivotDataFieldNaming extends PivotDataFieldPrefixes {
  /** Joins prefix and field name; `{prefix}` and `{field}` are substituted. */
  template: string;
}

export const ENGLISH_PIVOT_DATA_FIELD_NAMING: PivotDataFieldNaming = {
  template: '{prefix} {field}',
  sum: 'Sum of',
  count: 'Count of',
  average: 'Average of',
  max: 'Max of',
  min: 'Min of',
  product: 'Product of',
  countNumbers: 'Count of',
  stdDev: 'StdDev of',
  stdDevP: 'StdDevp of',
  variance: 'Var of',
  varianceP: 'Varp of',
};

const PREFIX_KEY: Record<PivotAggregation, keyof PivotDataFieldPrefixes> = {
  [PivotAggregation.Sum]: 'sum',
  [PivotAggregation.Count]: 'count',
  [PivotAggregation.Average]: 'average',
  [PivotAggregation.Max]: 'max',
  [PivotAggregation.Min]: 'min',
  [PivotAggregation.Product]: 'product',
  [PivotAggregation.CountNumbers]: 'countNumbers',
  [PivotAggregation.StdDev]: 'stdDev',
  [PivotAggregation.StdDevP]: 'stdDevP',
  [PivotAggregation.Var]: 'variance',
  [PivotAggregation.VarP]: 'varianceP',
};

/** Default value-field name, e.g. "Sum of Sales" or "合計 / Sales". */
export const pivotDataFieldName = (
  aggregation: PivotAggregation,
  field: string,
  naming: PivotDataFieldNaming = ENGLISH_PIVOT_DATA_FIELD_NAMING,
): string =>
  naming.template
    .replace('{prefix}', () => naming[PREFIX_KEY[aggregation]])
    .replace('{field}', () => field);
