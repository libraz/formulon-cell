import type {
  CategoryChartData,
  ChartData,
  ChartFilterSpec,
  ChartSpec,
  XYChartData,
} from './types.js';

/** Internal view of a category chart after applying the chart filter. */
export interface CategoryChartView {
  readonly data: CategoryChartData;
  readonly originalSeriesIndices: readonly number[];
  readonly originalCategoryIndices: readonly number[];
}

/** Internal view of an XY chart after applying the chart filter. */
export interface XYChartView {
  readonly data: XYChartData;
  readonly originalSeriesIndices: readonly number[];
  readonly originalPointIndices: readonly (readonly number[])[];
}

export type PreparedChartData = CategoryChartView | XYChartView;

const categoryIdAt = (data: CategoryChartData, index: number): string =>
  data.categoryIds?.[index] ?? `category-${index}`;

const pointIdAt = (point: XYChartData['series'][number]['points'][number], index: number): string =>
  point.id ?? `point-${index}`;

const hidden = (ids: readonly string[] | undefined, id: string): boolean =>
  ids?.includes(id) ?? false;

const prepareCategoryData = (
  data: CategoryChartData,
  filter: ChartFilterSpec | undefined,
): CategoryChartView => {
  const originalCategoryIndices = data.categories
    .map((_category, index) => index)
    .filter((index) => !hidden(filter?.hiddenCategoryIds, categoryIdAt(data, index)));
  const originalSeriesIndices = data.series
    .map((_series, index) => index)
    .filter((index) => {
      const series = data.series[index];
      return series !== undefined && !hidden(filter?.hiddenSeriesIds, series.id);
    });

  const series = originalSeriesIndices.map((originalIndex) => {
    const source = data.series[originalIndex];
    if (!source) return { id: `series-${originalIndex}`, values: [] as readonly number[] };
    return {
      id: source.id,
      ...(source.name === undefined ? {} : { name: source.name }),
      values: originalCategoryIndices.map((categoryIndex) => source.values[categoryIndex] ?? null),
    };
  });

  return {
    data: {
      domain: 'category',
      categories: originalCategoryIndices.map((index) => data.categories[index] ?? ''),
      categoryIds: originalCategoryIndices.map((index) => categoryIdAt(data, index)),
      series,
    },
    originalSeriesIndices,
    originalCategoryIndices,
  };
};

const prepareXYData = (data: XYChartData, filter: ChartFilterSpec | undefined): XYChartView => {
  const originalSeriesIndices = data.series
    .map((_series, index) => index)
    .filter((index) => {
      const series = data.series[index];
      return series !== undefined && !hidden(filter?.hiddenSeriesIds, series.id);
    });
  const originalPointIndices: number[][] = [];
  const series = originalSeriesIndices.map((originalIndex, visibleSeriesIndex) => {
    const source = data.series[originalIndex];
    if (!source) {
      originalPointIndices[visibleSeriesIndex] = [];
      return { id: `series-${originalIndex}`, points: [] };
    }
    const indexes = source.points
      .map((_point, index) => index)
      .filter((index) => {
        const point = source.points[index];
        return point !== undefined && !hidden(filter?.hiddenCategoryIds, pointIdAt(point, index));
      });
    originalPointIndices[visibleSeriesIndex] = indexes;
    return {
      id: source.id,
      ...(source.name === undefined ? {} : { name: source.name }),
      points: indexes.map((index) => {
        const point = source.points[index];
        if (!point) return { id: pointIdAt({ x: null, y: null }, index), x: null, y: null };
        return {
          ...(point.id === undefined ? { id: pointIdAt(point, index) } : { id: point.id }),
          x: point.x,
          y: point.y,
        };
      }),
    };
  });
  return {
    data: { domain: 'xy', series },
    originalSeriesIndices,
    originalPointIndices,
  };
};

/**
 * Apply chart filters to a cloned view. The source data and all nested arrays
 * remain untouched; the returned index maps let renderers retain original
 * identities and palette positions after compacting visible slots.
 */
export const prepareChartData = (data: ChartData, spec: ChartSpec): PreparedChartData =>
  data.domain === 'category'
    ? prepareCategoryData(data, spec.filter)
    : prepareXYData(data, spec.filter);
