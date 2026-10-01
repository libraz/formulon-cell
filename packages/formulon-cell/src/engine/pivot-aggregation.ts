import { PivotAggregation } from './types.js';

const PIVOT_AGGREGATION_NAMES: Record<PivotAggregation, string> = {
  [PivotAggregation.Sum]: 'Sum',
  [PivotAggregation.Count]: 'Count',
  [PivotAggregation.Average]: 'Average',
  [PivotAggregation.Max]: 'Max',
  [PivotAggregation.Min]: 'Min',
  [PivotAggregation.Product]: 'Product',
  [PivotAggregation.CountNumbers]: 'Count Numbers',
  [PivotAggregation.StdDev]: 'StdDev',
  [PivotAggregation.StdDevP]: 'StdDevP',
  [PivotAggregation.Var]: 'Var',
  [PivotAggregation.VarP]: 'VarP',
};

/** Stable display names for every aggregation ordinal accepted by formulon. */
export const pivotAggregationName = (aggregation: PivotAggregation): string =>
  PIVOT_AGGREGATION_NAMES[aggregation];
