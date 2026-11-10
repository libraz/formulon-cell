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
  data.categories[index]?.id ?? `category-${index}`;

const pointIdAt = (point: XYChartData['series'][number]['points'][number], index: number): string =>
  point.id ?? `point-${index}`;

const hidden = (ids: ReadonlySet<string> | undefined, id: string): boolean => ids?.has(id) ?? false;

const idSet = (ids: readonly string[] | undefined): ReadonlySet<string> | undefined =>
  ids === undefined ? undefined : new Set(ids);

const prepareCategoryData = (
  data: CategoryChartData,
  filter: ChartFilterSpec | undefined,
): CategoryChartView => {
  const hiddenCategories = idSet(filter?.hiddenCategoryIds);
  const hiddenSeries = idSet(filter?.hiddenSeriesIds);
  const originalCategoryIndices = data.categories
    .map((_category, index) => index)
    .filter((index) => !hidden(hiddenCategories, categoryIdAt(data, index)));
  const originalSeriesIndices = data.series
    .map((_series, index) => index)
    .filter((index) => {
      const series = data.series[index];
      return series !== undefined && !hidden(hiddenSeries, series.id);
    });

  const series = originalSeriesIndices.map((originalIndex) => {
    const source = data.series[originalIndex];
    if (!source) return { id: `series-${originalIndex}`, points: [] };
    return {
      id: source.id,
      ...(source.name === undefined ? {} : { name: source.name }),
      points: originalCategoryIndices.map(
        (categoryIndex) => source.points[categoryIndex] ?? { value: null },
      ),
    };
  });

  return {
    data: {
      domain: 'category',
      categories: originalCategoryIndices.map((index) => ({
        id: categoryIdAt(data, index),
        label: data.categories[index]?.label ?? '',
      })),
      series,
    },
    originalSeriesIndices,
    originalCategoryIndices,
  };
};

const prepareXYData = (data: XYChartData, filter: ChartFilterSpec | undefined): XYChartView => {
  const hiddenCategories = idSet(filter?.hiddenCategoryIds);
  const hiddenSeries = idSet(filter?.hiddenSeriesIds);
  const originalSeriesIndices = data.series
    .map((_series, index) => index)
    .filter((index) => {
      const series = data.series[index];
      return series !== undefined && !hidden(hiddenSeries, series.id);
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
        return point !== undefined && !hidden(hiddenCategories, pointIdAt(point, index));
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
