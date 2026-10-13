import type { ChartMarkerSpec, ChartSpec } from './types.js';

const cloneMarker = (marker: ChartMarkerSpec): ChartMarkerSpec => ({
  ...marker,
  ...(marker.fill ? { fill: { ...marker.fill } } : {}),
  ...(marker.line ? { line: { ...marker.line } } : {}),
});

/** Detach nested chart settings at store, history, and public API boundaries. */
export function cloneChartSpec<T extends ChartSpec>(value: T): T {
  return {
    ...value,
    ...(value.palette ? { palette: [...value.palette] } : {}),
    ...(value.legend ? { legend: { ...value.legend } } : {}),
    ...(value.xAxis ? { xAxis: { ...value.xAxis } } : {}),
    ...(value.yAxis ? { yAxis: { ...value.yAxis } } : {}),
    ...(value.filter
      ? {
          filter: {
            ...value.filter,
            ...(value.filter.hiddenSeriesIds
              ? { hiddenSeriesIds: [...value.filter.hiddenSeriesIds] }
              : {}),
            ...(value.filter.hiddenCategoryIds
              ? { hiddenCategoryIds: [...value.filter.hiddenCategoryIds] }
              : {}),
          },
        }
      : {}),
    ...(value.dataLabels ? { dataLabels: { ...value.dataLabels } } : {}),
    ...(value.seriesStyles
      ? {
          seriesStyles: value.seriesStyles.map((series) => ({
            ...series,
            ...(series.fill ? { fill: { ...series.fill } } : {}),
            ...(series.line ? { line: { ...series.line } } : {}),
            ...(series.marker ? { marker: cloneMarker(series.marker) } : {}),
            ...(series.points
              ? {
                  points: series.points.map((point) => ({
                    ...point,
                    ...(point.fill ? { fill: { ...point.fill } } : {}),
                    ...(point.line ? { line: { ...point.line } } : {}),
                    ...(point.marker ? { marker: cloneMarker(point.marker) } : {}),
                  })),
                }
              : {}),
          })),
        }
      : {}),
  };
}
