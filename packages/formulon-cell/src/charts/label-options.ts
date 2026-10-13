import type { ChartLabelPosition, ChartSpec } from './types.js';

type ChartLabelSpec = Pick<ChartSpec, 'kind' | 'grouping'>;

const PIE_LABEL_POSITIONS = ['best-fit', 'outside-end', 'inside-end', 'center'] as const;
const CLUSTERED_BAR_LABEL_POSITIONS = [
  'inside-base',
  'inside-end',
  'outside-end',
  'center',
] as const;
const STACKED_BAR_LABEL_POSITIONS = ['inside-base', 'inside-end', 'center'] as const;
const LINE_LABEL_POSITIONS = ['left', 'right', 'below', 'above', 'center'] as const;

/**
 * Return the positions supported by the corresponding Excel chart type.
 * Area charts intentionally expose no dLblPos values: Excel uses its
 * chart-specific center default when dLblPos is absent.
 */
export const chartLabelPositions = (spec: ChartLabelSpec): readonly ChartLabelPosition[] => {
  switch (spec.kind) {
    case 'pie':
      return PIE_LABEL_POSITIONS;
    case 'doughnut':
      return [];
    case 'column':
    case 'bar':
      return spec.grouping === 'stacked' || spec.grouping === 'percent-stacked'
        ? STACKED_BAR_LABEL_POSITIONS
        : CLUSTERED_BAR_LABEL_POSITIONS;
    case 'line':
    case 'scatter':
      return LINE_LABEL_POSITIONS;
    case 'area':
      return [];
  }
};

/** Excel's default data-label position for each chart family/grouping. */
export const defaultChartLabelPosition = (spec: ChartLabelSpec): ChartLabelPosition => {
  switch (spec.kind) {
    case 'pie':
      return 'best-fit';
    case 'doughnut':
      return 'center';
    case 'column':
    case 'bar':
      return spec.grouping === 'stacked' || spec.grouping === 'percent-stacked'
        ? 'center'
        : 'outside-end';
    case 'line':
    case 'scatter':
      return 'right';
    case 'area':
      return 'center';
  }
};

/**
 * Resolve the public convenience alias and unsupported positions before
 * chart-family geometry. `best-fit` is a real pie position; elsewhere it is
 * the API spelling for the chart's Excel default.
 */
export const resolveChartLabelPosition = (
  spec: ChartLabelSpec,
  requested?: ChartLabelPosition,
): ChartLabelPosition => {
  const fallback = defaultChartLabelPosition(spec);
  if (requested === undefined || requested === 'best-fit') return fallback;
  return chartLabelPositions(spec).includes(requested) ? requested : fallback;
};

export type { ChartLabelSpec };
