import type {
  ChartFillSpec,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartSeriesStyleSpec,
  ChartSpec,
} from './types.js';

export const DEFAULT_PALETTE = [
  '#4472c4',
  '#ed7d31',
  '#70ad47',
  '#ffc000',
  '#5b9bd5',
  '#a5a5a5',
  '#264478',
  '#9e480e',
] as const;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

const transparencyOpacity = (transparency: number | undefined): number | undefined =>
  transparency === undefined || !Number.isFinite(transparency)
    ? undefined
    : (100 - clamp(transparency, 0, 100)) / 100;

const mergeFill = (base: ChartFillSpec | undefined, override: ChartFillSpec | undefined) =>
  base === undefined && override === undefined
    ? undefined
    : {
        ...(base ?? {}),
        ...(override ?? {}),
      };

const mergeLine = (base: ChartLineSpec | undefined, override: ChartLineSpec | undefined) =>
  base === undefined && override === undefined
    ? undefined
    : {
        ...(base ?? {}),
        ...(override ?? {}),
      };

export interface ResolvedChartStyle {
  readonly color: string;
  readonly fill: ChartFillSpec | undefined;
  readonly line: ChartLineSpec | undefined;
  readonly marker: ChartMarkerSpec;
}

const styleFor = (spec: ChartSpec, seriesId: string): ChartSeriesStyleSpec | undefined =>
  spec.seriesStyles?.find((style) => style.seriesId === seriesId);

/** Resolve point overrides field-by-field so a point can change only its fill color. */
export const resolveChartStyle = (
  spec: ChartSpec,
  palette: readonly string[],
  seriesId: string,
  pointId: string | undefined,
  originalSeriesIndex: number,
  originalPointIndex: number | undefined,
): ResolvedChartStyle => {
  const seriesStyle = styleFor(spec, seriesId);
  const pointStyle =
    pointId === undefined
      ? undefined
      : seriesStyle?.points?.find((style) => style.pointId === pointId);
  const varyColors = spec.varyColors ?? (spec.kind === 'pie' || spec.kind === 'doughnut');
  const paletteIndex = varyColors
    ? (originalPointIndex ?? originalSeriesIndex)
    : originalSeriesIndex;
  const color =
    palette[((paletteIndex % palette.length) + palette.length) % palette.length] ?? '#4472c4';
  const markerBase = seriesStyle?.marker;
  const markerOverride = pointStyle?.marker;
  const defaultMarkerVisible = !(
    spec.kind === 'scatter' && (spec.scatterStyle ?? 'marker') === 'line'
  );
  const marker: ChartMarkerSpec = {
    // Scatter line charts hide their default markers while preserving an
    // explicit series/point visibility setting for the settings UI and
    // renderer. Every other chart keeps the historical visible default.
    ...(markerBase ?? {}),
    ...(markerOverride ?? {}),
    // Treat an explicitly present `undefined` as omitted so a partial style
    // object cannot accidentally turn the line-only default back on.
    visible: markerOverride?.visible ?? markerBase?.visible ?? defaultMarkerVisible,
    // A point-level general fill/line is inherited by its marker too. An
    // explicit point marker fill/line still wins over that inherited value.
    fill: mergeFill(mergeFill(markerBase?.fill, pointStyle?.fill), markerOverride?.fill),
    line: mergeLine(mergeLine(markerBase?.line, pointStyle?.line), markerOverride?.line),
  };
  return {
    color,
    fill: mergeFill(seriesStyle?.fill, pointStyle?.fill),
    line: mergeLine(seriesStyle?.line, pointStyle?.line),
    marker,
  };
};

export const fillColor = (fill: ChartFillSpec | undefined, fallback: string): string =>
  fill?.color ?? fallback;

export const fillOpacity = (
  fill: ChartFillSpec | undefined,
  fallback?: number,
): number | undefined => transparencyOpacity(fill?.transparency) ?? fallback;

export const lineOpacity = (line: ChartLineSpec | undefined): number | undefined =>
  transparencyOpacity(line?.transparency);

export const lineWidth = (
  line: ChartLineSpec | undefined,
  fallback?: number,
): number | undefined =>
  line?.width !== undefined && Number.isFinite(line.width) ? clamp(line.width, 0, 72) : fallback;

export const dashArray = (dash: ChartLineSpec['dash']): string | undefined => {
  switch (dash) {
    case 'dash':
      return '6 4';
    case 'dot':
      return '1 4';
    case 'dash-dot':
      return '6 3 1 3';
    default:
      return undefined;
  }
};

export const markerSize = (marker: ChartMarkerSpec): number =>
  marker.size !== undefined && Number.isFinite(marker.size) ? clamp(marker.size, 2, 72) : 5;

export const markerVisible = (marker: ChartMarkerSpec): boolean => marker.visible !== false;
