import { shapeDashStroke } from '../commands/shape-format.js';
import { clamp } from './numeric.js';
import { DEFAULT_THEME_ACCENTS } from './series-colors.js';
import type {
  ChartAreaSpec,
  ChartFillSpec,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartSeriesStyleSpec,
  ChartSpec,
} from './types.js';

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

export const styleFor = (spec: ChartSpec, seriesId: string): ChartSeriesStyleSpec | undefined =>
  spec.seriesStyles?.find((style) => style.seriesId === seriesId);

const colorOnly = (color: string | undefined): ChartFillSpec | undefined =>
  color === undefined ? undefined : { color };

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
    palette[((paletteIndex % palette.length) + palette.length) % palette.length] ??
    DEFAULT_THEME_ACCENTS[0] ??
    '#156082';
  // Excel keeps no fill on a line or scatter series or point: a fill colour stands in
  // for the palette entry on the line and markers, and its transparency is not drawn.
  const stroked = spec.kind === 'line' || spec.kind === 'scatter';
  const pointColor = stroked ? colorOnly(pointStyle?.fill?.color) : undefined;
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
    fill: mergeFill(
      mergeFill(markerBase?.fill, stroked ? pointColor : pointStyle?.fill),
      markerOverride?.fill,
    ),
    line: mergeLine(
      mergeLine(mergeLine(markerBase?.line, pointStyle?.line), pointColor),
      markerOverride?.line,
    ),
  };
  if (stroked) {
    return {
      color: pointColor?.color ?? seriesStyle?.fill?.color ?? color,
      fill: undefined,
      line: mergeLine(
        seriesStyle?.line,
        pointStyle?.line?.color === undefined
          ? mergeLine(pointStyle?.line, pointColor)
          : pointStyle.line,
      ),
      marker,
    };
  }
  return {
    color,
    fill: mergeFill(seriesStyle?.fill, pointStyle?.fill),
    line: mergeLine(seriesStyle?.line, pointStyle?.line),
    marker,
  };
};

/** Excel's chart-area outline where the file states none. */
export const DEFAULT_CHART_AREA_LINE = { color: '#898989', widthPt: 1 } as const;

/** The chart area's outline as the file states it, over the default. A stated No Line is null. */
export const chartAreaLine = (
  area: ChartAreaSpec | undefined,
): {
  color: string;
  widthPt: number;
  /** The drawn width in scene pixels. */
  widthPx: number;
  opacity?: number;
  dash?: ChartLineSpec['dash'];
} | null => {
  if (area?.line === 'none') return null;
  const opacity = lineOpacity(area?.line);
  const stated = area?.line?.width;
  const widthPt =
    stated !== undefined && Number.isFinite(stated)
      ? clamp(stated, 0, 72)
      : DEFAULT_CHART_AREA_LINE.widthPt;
  return {
    color: area?.line?.color ?? DEFAULT_CHART_AREA_LINE.color,
    widthPt,
    widthPx: widthPt * PX_PER_PT,
    ...(opacity === undefined ? {} : { opacity }),
    ...(area?.line?.dash === undefined ? {} : { dash: area.line.dash }),
  };
};

/** The chart area's fill as the file states it: undefined when it states none of its own (the
 *  chart is filled with the surface), null for a stated No Fill. */
export const chartAreaFill = (
  area: ChartAreaSpec | undefined,
): { color: string; opacity: number } | null | undefined => {
  if (area?.fill === 'none') return null;
  if (area?.fill?.color === undefined) return undefined;
  return { color: area.fill.color, opacity: fillOpacity(area.fill, 1) ?? 1 };
};

export const fillColor = (fill: ChartFillSpec | undefined, fallback: string): string =>
  fill?.color ?? fallback;

export const fillOpacity = (
  fill: ChartFillSpec | undefined,
  fallback?: number,
): number | undefined => transparencyOpacity(fill?.transparency) ?? fallback;

export const lineOpacity = (line: ChartLineSpec | undefined): number | undefined =>
  transparencyOpacity(line?.transparency);

/** CSS pixels per point: a scene is laid out in pixels and a stated line width is in points. */
export const PX_PER_PT = 96 / 72;

/** The drawn width of a line or scatter series line that states none: the spreadsheet reads it as 3 pt. */
export const DEFAULT_SERIES_LINE_PX = 3 * PX_PER_PT;

/** The drawn stroke width in scene pixels of a line stating its width in points, or `fallback` (pixels). */
export const lineWidth = (
  line: ChartLineSpec | undefined,
  fallback?: number,
): number | undefined =>
  line?.width !== undefined && Number.isFinite(line.width)
    ? clamp(line.width, 0, 72) * PX_PER_PT
    : fallback;

/** The SVG stroke attributes of a dash style at a drawn `width` in scene pixels; none for a solid line. */
export const dashAttrs = (dash: ChartLineSpec['dash'], width: number): Record<string, string> => {
  if (dash === undefined || dash === 'solid') return {};
  const stroke = shapeDashStroke(dash, width);
  return {
    'stroke-dasharray': stroke.dasharray,
    ...(stroke.linecap === 'round' ? { 'stroke-linecap': 'round' } : {}),
  };
};

export const markerSize = (marker: ChartMarkerSpec): number =>
  marker.size !== undefined && Number.isFinite(marker.size) ? clamp(marker.size, 2, 72) : 5;

export const markerVisible = (marker: ChartMarkerSpec): boolean => marker.visible !== false;

/** A series' key colour: a line or scatter series shows its line, or its markers when it has none. */
export const keyColor = (spec: ChartSpec, style: ReturnType<typeof resolveChartStyle>): string => {
  if (spec.kind !== 'line' && spec.kind !== 'scatter') return fillColor(style.fill, style.color);
  const lineless = spec.kind === 'scatter' && (spec.scatterStyle ?? 'marker') === 'marker';
  return (lineless ? style.marker.fill?.color : style.line?.color) ?? style.color;
};
