import type { CategoryChartView, PreparedChartData, XYChartView } from './filters.js';
import { prepareChartData } from './filters.js';
import { formatChartValue } from './format.js';
import { resolveChartLabelPosition } from './label-options.js';
import {
  automaticAxisDomain,
  formatTick,
  linearScale,
  majorTicks,
  niceDomain,
  powerTicks,
  resolvePositivePowerDomain,
  unclampedLogProjection,
  validateLogBase,
} from './scales.js';
import { type ScatterLineSample, scatterLineSegments } from './scatter-lines.js';
import { automaticPalette } from './series-colors.js';
import {
  dashArray,
  fillColor,
  fillOpacity,
  lineOpacity,
  lineWidth,
  markerSize,
  markerVisible,
  resolveChartStyle,
} from './styles.js';
import type {
  CategoryChartData,
  ChartAxisSpec,
  ChartData,
  ChartGrouping,
  ChartKind,
  ChartPrimitive,
  ChartScene,
  ChartSpec,
  ChartValueFormatter,
  XYChartData,
} from './types.js';

/** Excel sets a chart title at 14 pt over 9 or 10 pt chart text. */
const TITLE_FONT_SIZE = 14;
const TITLE_BASELINE = 16;
/** The top margin of a plot below a title. */
const TITLE_RESERVE = 30;

const finite = (value: number, fallback: number): number =>
  Number.isFinite(value) ? value : fallback;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

const isCircularKind = (kind: ChartKind): boolean => kind === 'pie' || kind === 'doughnut';

const textWidth = (text: string, fontSize = 10): number => {
  const scale = finite(fontSize, 10) / 10;
  let width = 0;
  for (const character of Array.from(text)) {
    const code = character.codePointAt(0) ?? 0;
    const wide =
      (code >= 0x1100 && code <= 0x11ff) ||
      (code >= 0x2e80 && code <= 0xa4cf) ||
      (code >= 0xac00 && code <= 0xd7af) ||
      (code >= 0xf900 && code <= 0xfaff) ||
      (code >= 0xff01 && code <= 0xff60) ||
      code >= 0x1f300;
    const characterWidth =
      code === 0x25a0 ? 9.2 : code === 0x2026 ? 8.2 : wide ? 10 : character === ' ' ? 3.5 : 5.8;
    width += characterWidth;
  }
  return Math.max(fontSize === 10 ? 12 : Math.max(1, fontSize * 0.4), width * scale);
};

const truncateLabel = (label: string, maxWidth: number, fontSize = 10): string => {
  if (textWidth(label, fontSize) <= maxWidth) return label;
  const suffix = '…';
  let result = '';
  for (const character of Array.from(label)) {
    const candidate = `${result}${character}${suffix}`;
    if (textWidth(candidate, fontSize) > maxWidth) break;
    result += character;
  }
  return result.length > 0 ? `${result}${suffix}` : suffix;
};

const primitive = (
  tag: ChartPrimitive['tag'],
  attrs: Record<string, string | number>,
  role?: ChartPrimitive['role'],
  seriesIndex?: number,
  pointIndex?: number,
  text?: string,
  identity?: { readonly seriesId?: string; readonly pointId?: string },
): ChartPrimitive => ({
  tag,
  attrs:
    tag === 'text'
      ? {
          fill: 'currentColor',
          'font-family': 'system-ui, sans-serif',
          'font-size': role === 'title' ? TITLE_FONT_SIZE : 10,
          ...attrs,
        }
      : attrs,
  ...(role === undefined ? {} : { role }),
  ...(seriesIndex === undefined ? {} : { seriesIndex }),
  ...(pointIndex === undefined ? {} : { pointIndex }),
  ...(identity?.seriesId === undefined ? {} : { seriesId: identity.seriesId }),
  ...(identity?.pointId === undefined ? {} : { pointId: identity.pointId }),
  ...(text === undefined ? {} : { text }),
});

/** Re-attach the engine's per-point states and series formats, which the filtered view does not carry. */
const withPointStates = (view: PreparedChartData, source: ChartData): PreparedChartData => {
  if (view.data.domain === 'category' && source.domain === 'category') {
    const categoryView = view as CategoryChartView;
    return {
      ...categoryView,
      data: {
        ...categoryView.data,
        series: categoryView.data.series.map((series, index) => {
          const original = source.series[categoryView.originalSeriesIndices[index] ?? index];
          const states = original?.states;
          return {
            ...series,
            ...(original?.formatCode === undefined ? {} : { formatCode: original.formatCode }),
            ...(states
              ? { states: categoryView.originalCategoryIndices.map((i) => states[i] ?? 'value') }
              : {}),
          };
        }),
      },
    };
  }
  if (view.data.domain === 'xy' && source.domain === 'xy') {
    const xyView = view as XYChartView;
    return {
      ...xyView,
      data: {
        ...xyView.data,
        series: xyView.data.series.map((series, seriesIndex) => {
          const original =
            source.series[xyView.originalSeriesIndices[seriesIndex] ?? seriesIndex]?.points;
          const sourceSeries =
            source.series[xyView.originalSeriesIndices[seriesIndex] ?? seriesIndex];
          return {
            ...series,
            ...(sourceSeries?.formatCode === undefined
              ? {}
              : { formatCode: sourceSeries.formatCode }),
            ...(sourceSeries?.xFormatCode === undefined
              ? {}
              : { xFormatCode: sourceSeries.xFormatCode }),
            points: series.points.map((point, pointIndex) => {
              const state =
                original?.[xyView.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex]
                  ?.state;
              return state === undefined ? point : { ...point, state };
            }),
          };
        }),
      },
    };
  }
  return view;
};

/** Series-level labels override the chart-wide labels; hidden series labels stay hidden. */
const labelSpecFor = (
  spec: ChartSpec,
  series: { readonly id: string; readonly formatCode?: string },
): ChartSpec['dataLabels'] => {
  const labels =
    spec.seriesStyles?.find((style) => style.seriesId === series.id)?.dataLabels ?? spec.dataLabels;
  // The percentage part derives its own code from `numberFormat`, so only value labels inherit.
  if (!labels || labels.showPercentage || !series.formatCode) return labels;
  if (labels.numberFormat !== undefined && labels.numberFormat.trim().length > 0) return labels;
  return { ...labels, numberFormat: series.formatCode };
};

const hasFormat = (code: string | undefined): code is string =>
  code !== undefined && code.trim().length > 0 && code.trim().toLowerCase() !== 'general';

/** Format numeric category labels with the engine's category format code. */
const withCategoryFormat = (
  view: PreparedChartData,
  source: ChartData,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): PreparedChartData => {
  if (view.data.domain !== 'category' || source.domain !== 'category') return view;
  const code = source.categoryFormatCode;
  if (!hasFormat(code)) return view;
  const categoryView = view as CategoryChartView;
  return {
    ...categoryView,
    data: {
      ...categoryView.data,
      categories: categoryView.data.categories.map((label) => {
        const value = label.trim().length > 0 ? Number(label) : Number.NaN;
        return Number.isFinite(value)
          ? (formatChartValue(value, code, locale, formatter) ?? label)
          : label;
      }),
    },
  };
};

/** Default the value-axis format to the first plotted series' format code. */
const withValueAxisFormat = (spec: ChartSpec, view: PreparedChartData): ChartSpec => {
  if (view.data.domain === 'xy') {
    // A scatter chart's horizontal axis follows the x values' format, its vertical one the y values'.
    const first = view.data.series[0];
    const linked = (axis: ChartAxisSpec | undefined, code: string | undefined) =>
      hasFormat(code) && !((axis?.numberFormat?.trim().length ?? 0) > 0)
        ? { ...axis, numberFormat: code }
        : axis;
    const xAxis = linked(spec.xAxis, first?.xFormatCode);
    const yAxis = linked(spec.yAxis, first?.formatCode);
    return {
      ...spec,
      ...(xAxis === undefined ? {} : { xAxis }),
      ...(yAxis === undefined ? {} : { yAxis }),
    };
  }
  if (view.data.domain !== 'category' || isCircularKind(spec.kind)) return spec;
  if (spec.grouping === 'percent-stacked') return spec;
  const code = view.data.series[0]?.formatCode;
  if (!hasFormat(code)) return spec;
  const key = spec.kind === 'bar' ? 'xAxis' : 'yAxis';
  const axis = spec[key];
  if ((axis?.numberFormat?.trim().length ?? 0) > 0) return spec;
  return { ...spec, [key]: { ...axis, numberFormat: code } };
};

const hasDataLabels = (spec: ChartSpec): boolean =>
  [spec.dataLabels, ...(spec.seriesStyles?.map((style) => style.dataLabels) ?? [])].some(
    (labels) => labels !== undefined && labels.visible !== false,
  );

const dataHasValues = (
  data: ChartData,
  kind: ChartKind,
  numericLogBase?: number,
  scatterLogBases?: { readonly x?: number; readonly y?: number },
): boolean => {
  if (data.domain === 'category') {
    const valueIsPlottable = (value: number | null): boolean =>
      value !== null && Number.isFinite(value) && (numericLogBase === undefined || value > 0);
    if (kind === 'pie') {
      const first = data.series[0];
      return (
        first?.values.some((value) => value !== null && Number.isFinite(value) && value !== 0) ??
        false
      );
    }
    if (kind === 'doughnut') {
      return data.series.some((series) =>
        series.values.some((value) => value !== null && Number.isFinite(value) && value !== 0),
      );
    }
    return data.series.some((series) => series.values.some(valueIsPlottable));
  }
  return data.series.some((series) =>
    series.points.some((point) => {
      const { x, y } = point;
      return (
        x !== null &&
        y !== null &&
        Number.isFinite(x) &&
        Number.isFinite(y) &&
        (scatterLogBases?.x === undefined || x > 0) &&
        (scatterLogBases?.y === undefined || y > 0)
      );
    }),
  );
};

const valueFor = (value: number | null): number | null =>
  value !== null && Number.isFinite(value) ? value : null;

const boundedOption = (
  value: number | undefined,
  min: number,
  max: number,
  fallback: number,
): number => (value !== undefined && Number.isFinite(value) ? clamp(value, min, max) : fallback);

const saturatedAdd = (left: number, right: number): number => {
  const sum = left + right;
  if (Number.isFinite(sum)) return sum;
  return left >= 0 && right >= 0 ? Number.MAX_VALUE : -Number.MAX_VALUE;
};

/** The spec's palette, else Excel's automatic colours in the workbook theme. */
const paletteFor = (spec: ChartSpec, accents: readonly string[] | undefined): readonly string[] =>
  spec.palette && spec.palette.length > 0 ? spec.palette : automaticPalette(accents);

const colorFor = (palette: readonly string[], index: number): string =>
  palette[((index % palette.length) + palette.length) % palette.length] ?? '#156082';

/** A series' key colour: a line or scatter series shows its line, or its markers when it has none. */
const keyColor = (spec: ChartSpec, style: ReturnType<typeof resolveChartStyle>): string => {
  if (spec.kind !== 'line' && spec.kind !== 'scatter') return fillColor(style.fill, style.color);
  const lineless = spec.kind === 'scatter' && (spec.scatterStyle ?? 'marker') === 'marker';
  return (lineless ? style.marker.fill?.color : style.line?.color) ?? style.color;
};

/**
 * The text a chart title shows: its own text, nothing for `''`, and for Excel's
 * automatic title the name of a chart's only series, else `chartTitle`.
 */
export const chartTitleText = (
  title: ChartSpec['title'],
  data: ChartData,
  chartTitle = 'Chart Title',
): string | undefined => {
  if (title === undefined) return undefined;
  if (typeof title === 'string') return title.length > 0 ? title : undefined;
  const only = data.series.length === 1 ? data.series[0]?.name : undefined;
  return only !== undefined && only.length > 0 ? only : chartTitle;
};

/** The text an axis title shows: its own text, or `axisTitle` for Excel's automatic title. */
export const axisTitleText = (
  title: ChartAxisSpec['title'],
  axisTitle = 'Axis Title',
): string | undefined => {
  if (title === undefined) return undefined;
  if (typeof title === 'string') return title.length > 0 ? title : undefined;
  return axisTitle;
};

/** Whether a line or scatter series draws Excel's smoothed line. */
const smoothed = (spec: ChartSpec, seriesId: string): boolean =>
  (spec.kind === 'line' || spec.kind === 'scatter') &&
  spec.seriesStyles?.find((style) => style.seriesId === seriesId)?.smooth === true;

/**
 * Excel's smoothed line through `points`: cubic Béziers from Catmull-Rom with
 * tension 1/6, the end points repeated as their own neighbours.
 */
const smoothPath = (points: readonly (readonly [number, number])[]): string => {
  const commands = points.slice(0, 1).map(([x, y]) => `M${x},${y}`);
  for (let index = 0; index + 1 < points.length; index += 1) {
    const p1 = points[index] as readonly [number, number];
    const p2 = points[index + 1] as readonly [number, number];
    const p0 = points[index - 1] ?? p1;
    const p3 = points[index + 2] ?? p2;
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    commands.push(`C${c1[0]},${c1[1]} ${c2[0]},${c2[1]} ${p2[0]},${p2[1]}`);
  }
  return commands.join(' ');
};

const categoryIdAt = (view: CategoryChartView, pointIndex: number): string =>
  view.data.categoryIds?.[pointIndex] ??
  `category-${view.originalCategoryIndices[pointIndex] ?? pointIndex}`;

const pointIdAt = (view: XYChartView, seriesIndex: number, pointIndex: number): string => {
  const point = view.data.series[seriesIndex]?.points[pointIndex];
  return point?.id ?? `point-${view.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex}`;
};

const styleAttrs = (
  attrs: Record<string, string | number>,
  style: ReturnType<typeof resolveChartStyle>,
  fallbackOpacity?: number,
): Record<string, string | number> => {
  attrs.fill = fillColor(style.fill, style.color);
  const opacity = fillOpacity(style.fill, fallbackOpacity);
  if (opacity !== undefined) attrs['fill-opacity'] = opacity;
  if (style.line) {
    attrs.stroke = style.line.color ?? style.color;
    const strokeOpacity = lineOpacity(style.line);
    if (strokeOpacity !== undefined) attrs['stroke-opacity'] = strokeOpacity;
    const width = lineWidth(style.line);
    if (width !== undefined) attrs['stroke-width'] = width;
    const dash = dashArray(style.line.dash);
    if (dash !== undefined) attrs['stroke-dasharray'] = dash;
  }
  return attrs;
};

const markerAttrs = (
  x: number,
  y: number,
  style: ReturnType<typeof resolveChartStyle>,
): { tag: ChartPrimitive['tag']; attrs: Record<string, string | number> } => {
  const marker = style.marker;
  const size = markerSize(marker);
  const half = size / 2;
  const fill = fillColor(marker.fill ?? style.fill, style.color);
  const opacity = fillOpacity(marker.fill ?? style.fill);
  const line = marker.line;
  const stroke = line === undefined ? undefined : (line.color ?? style.color);
  const strokeOpacity = lineOpacity(line);
  const strokeWidth = lineWidth(line);
  const dash = dashArray(line?.dash);
  const common: Record<string, string | number> = {
    fill,
    ...(opacity === undefined ? {} : { 'fill-opacity': opacity }),
    ...(stroke === undefined ? {} : { stroke }),
    ...(strokeOpacity === undefined ? {} : { 'stroke-opacity': strokeOpacity }),
    ...(strokeWidth === undefined ? {} : { 'stroke-width': strokeWidth }),
    ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
  };
  // Excel's x, star and plus are drawn by their outline alone, in the outline colour.
  const strokedShape = (d: string) => ({
    tag: 'path' as const,
    attrs: {
      d,
      fill: 'none',
      stroke: stroke ?? style.color,
      'stroke-width': strokeWidth ?? 1,
      ...(strokeOpacity === undefined ? {} : { 'stroke-opacity': strokeOpacity }),
      ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
    },
  });
  const cross = `M${x - half},${y - half} L${x + half},${y + half} M${x - half},${y + half} L${x + half},${y - half}`;
  const upright = `M${x},${y - half} L${x},${y + half}`;
  const bar = (width: number) => ({
    tag: 'rect' as const,
    attrs: { x: x - width / 2, y: y - size / 10, width, height: size / 5, ...common },
  });
  switch (marker.shape ?? 'circle') {
    case 'x':
      return strokedShape(cross);
    case 'star':
      return strokedShape(`${cross} ${upright}`);
    case 'plus':
      return strokedShape(`M${x - half},${y} L${x + half},${y} ${upright}`);
    case 'dot':
      return bar(size / 2);
    case 'dash':
      return bar(size);
    case 'square':
      return {
        tag: 'rect',
        attrs: { x: x - half, y: y - half, width: size, height: size, ...common },
      };
    case 'diamond':
      return {
        tag: 'polygon',
        attrs: {
          points: `${x},${y - half} ${x + half},${y} ${x},${y + half} ${x - half},${y}`,
          ...common,
        },
      };
    case 'triangle':
      return {
        tag: 'polygon',
        attrs: {
          points: `${x},${y - half} ${x + half},${y + half} ${x - half},${y + half}`,
          ...common,
        },
      };
    default:
      // A circle, and a picture marker drawn as a plain one in the series colour.
      return { tag: 'circle', attrs: { cx: x, cy: y, r: half, ...common } };
  }
};

interface Domain {
  min: number;
  max: number;
  /** The major unit an automatic axis resolved along with its bounds. */
  majorUnit?: number;
}

const axisDomain = (
  values: readonly number[],
  explicitMin: number | undefined,
  explicitMax: number | undefined,
  percent: boolean,
  logBase?: number,
): Domain => {
  if (logBase !== undefined) {
    return resolvePositivePowerDomain(values, explicitMin, explicitMax, logBase);
  }
  const finiteValues = values.filter(Number.isFinite);
  let min = 0;
  let max = 1;
  if (finiteValues.length > 0) {
    min = finiteValues.reduce(
      (current, value) => Math.min(current, value),
      Number.POSITIVE_INFINITY,
    );
    max = finiteValues.reduce(
      (current, value) => Math.max(current, value),
      Number.NEGATIVE_INFINITY,
    );
  }
  if (percent) {
    if (min < 0 && max > 0) {
      min = -1;
      max = 1;
    } else if (max > 0) {
      min = 0;
      max = 1;
    } else {
      min = -1;
      max = 0;
    }
  }
  const hasExplicitMin = explicitMin !== undefined && Number.isFinite(explicitMin);
  const hasExplicitMax = explicitMax !== undefined && Number.isFinite(explicitMax);
  const baseMin = min;
  const baseMax = max;
  const spanForMissingSide = Math.max(
    Number.isFinite(baseMax - baseMin) && baseMax > baseMin ? baseMax - baseMin : 0,
    Math.abs(baseMin) * 0.1,
    Math.abs(baseMax) * 0.1,
    1,
  );
  if (hasExplicitMin && hasExplicitMax) {
    min = explicitMin as number;
    max = explicitMax as number;
    if (min > max) [min, max] = [max, min];
  } else if (hasExplicitMin) {
    min = explicitMin as number;
    if (max < min) max = Math.min(Number.MAX_VALUE, min + spanForMissingSide);
  } else if (hasExplicitMax) {
    max = explicitMax as number;
    if (min > max) min = Math.max(-Number.MAX_VALUE, max - spanForMissingSide);
  }
  if (!percent && (!hasExplicitMin || !hasExplicitMax)) {
    return automaticAxisDomain(min, max, {
      ...(hasExplicitMin ? { min } : {}),
      ...(hasExplicitMax ? { max } : {}),
    });
  }
  if (min === max) {
    const padding = Math.max(Math.abs(min) * 0.1, Number.MIN_VALUE, percent ? 1 : 1);
    if (hasExplicitMin && !hasExplicitMax) max = Math.min(Number.MAX_VALUE, min + padding);
    else if (hasExplicitMax && !hasExplicitMin) min = Math.max(-Number.MAX_VALUE, max - padding);
    else {
      min = Math.max(-Number.MAX_VALUE, min - padding);
      max = Math.min(Number.MAX_VALUE, max + padding);
    }
  }
  // Treat non-finite explicit bounds exactly like omitted bounds.  Checking
  // the raw option for `undefined` here would leave NaN/Infinity in the
  // auto-expansion decision even though the finite flags above rejected them.
  if (!hasExplicitMin || !hasExplicitMax) {
    const expanded = niceDomain(min, max, 4);
    if (!hasExplicitMin) min = expanded.min;
    if (!hasExplicitMax) max = expanded.max;
  }
  return { min, max };
};

const reversedAxis = (axis: ChartAxisSpec | undefined): boolean => axis?.reverseOrder === true;

const resolvedLogBase = (axis: ChartAxisSpec | undefined): number | undefined =>
  validateLogBase(axis?.logBase);

/** Numeric axes are log-capable only for the chart families with one. */
const categoryNumericLogBase = (spec: ChartSpec, axis: 'x' | 'y'): number | undefined => {
  if ((spec.grouping ?? 'clustered') !== 'clustered') return undefined;
  if (axis === 'x' && spec.kind === 'bar') return resolvedLogBase(spec.xAxis);
  if (axis === 'y' && (spec.kind === 'column' || spec.kind === 'line' || spec.kind === 'area')) {
    return resolvedLogBase(spec.yAxis);
  }
  return undefined;
};

const scatterLogBases = (spec: ChartSpec): { readonly x?: number; readonly y?: number } =>
  spec.kind === 'scatter'
    ? {
        x: resolvedLogBase(spec.xAxis),
        y: resolvedLogBase(spec.yAxis),
      }
    : {};

type ResolvedAxisLabelPosition = 'nextTo' | 'low' | 'high' | 'none';

/** Invalid runtime values retain the legacy low-edge label placement. */
const resolvedAxisLabelPosition = (axis: ChartAxisSpec | undefined): ResolvedAxisLabelPosition => {
  const position = axis?.labelPosition;
  return position === 'nextTo' || position === 'low' || position === 'high' || position === 'none'
    ? position
    : 'low';
};

const categoryIndex = (count: number, index: number, reverse: boolean): number =>
  reverse ? count - 1 - index : index;

const xCategory = (
  plot: ChartScene['plot'],
  count: number,
  index: number,
  reverse = false,
): number => {
  if (count <= 1) return plot.x + plot.width / 2;
  return plot.x + ((categoryIndex(count, index, reverse) + 0.5) * plot.width) / count;
};

const yCategory = (
  plot: ChartScene['plot'],
  count: number,
  index: number,
  reverse = false,
): number => {
  if (count <= 1) return plot.y + plot.height / 2;
  return plot.y + ((categoryIndex(count, index, reverse) + 0.5) * plot.height) / count;
};

/** Project an ascending numeric domain onto a screen range, optionally reversing its endpoints. */
const numericProjection = (
  value: number,
  domain: Domain,
  start: number,
  end: number,
  reverse: boolean,
  logBase?: number,
): number => {
  const from = reverse ? end : start;
  const to = reverse ? start : end;
  const projected =
    logBase === undefined
      ? (linearScale(value, domain.min, domain.max, from, to) ?? from)
      : unclampedLogProjection(value, domain.min, domain.max, from, to);
  return clamp(projected, Math.min(start, end), Math.max(start, end));
};

/**
 * Resolve a numeric crossing through the same clamped projection used by the
 * marks. Invalid runtime values (including non-finite numbers and unknown
 * strings) retain the historic zero crossing.
 */
const numericAxisCrossing = (
  axis: ChartAxisSpec | undefined,
  domain: Domain,
  scale: (value: number) => number,
  logBase?: number,
): number => {
  const crossing = axis?.crossing;
  if (crossing === 'min') return scale(domain.min);
  if (crossing === 'max') return scale(domain.max);
  if (
    typeof crossing === 'number' &&
    Number.isFinite(crossing) &&
    (logBase === undefined || crossing > 0)
  )
    return scale(crossing);
  if (logBase !== undefined) return scale(clamp(1, domain.min, domain.max));
  return scale(0);
};

/**
 * Resolve a category-axis crossing between bands. Explicit positions are in
 * logical category order and are reflected by reverseOrder. Automatic mode
 * deliberately keeps the legacy physical edge, independent of reversal.
 */
const categoryAxisCrossing = (
  axis: ChartAxisSpec | undefined,
  count: number,
  start: number,
  end: number,
  reverse: boolean,
  automatic: number,
): number => {
  if (count <= 0) return automatic;
  const crossing = axis?.crossing;
  if (crossing === undefined || crossing === 'auto') return automatic;
  let edge: number;
  if (crossing === 'min') edge = 0;
  else if (crossing === 'max') edge = count;
  else if (typeof crossing === 'number' && Number.isFinite(crossing)) {
    const categoryNumber = clamp(Math.round(crossing), 1, count);
    edge = categoryNumber - 1;
  } else {
    return automatic;
  }
  const slot = (end - start) / count;
  const reflectedEdge = reverse ? count - edge : edge;
  return start + reflectedEdge * slot;
};

const axisIsVisible = (axis: ChartAxisSpec | undefined): boolean => axis?.visible !== false;

const axisHasGridlines = (axis: ChartAxisSpec | undefined): boolean =>
  axis?.majorGridlines !== false;

const axisTicks = (
  domain: Domain,
  axis: ChartAxisSpec | undefined,
  logBase?: number,
): readonly number[] =>
  logBase === undefined
    ? majorTicks(domain.min, domain.max, axis?.majorUnit ?? domain.majorUnit)
    : powerTicks(domain.min, domain.max, logBase);

const axisTickLabel = (
  value: number,
  percent: boolean,
  axis: ChartAxisSpec | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): string => {
  const formatCode = axis?.numberFormat;
  if (formatCode !== undefined && formatCode.trim().length > 0) {
    return formatChartValue(value, formatCode, locale, formatter) ?? formatTick(value, percent);
  }
  return formatTick(value, percent);
};

const AXIS_LABEL_MAX_WIDTH = 160;
const AXIS_LABEL_MAX_CHARS = 256;
// Rotated SVG text can extend a few pixels to the left of its x coordinate
// under the default system font. Keep the baseline far enough inside the
// canvas for narrow charts while leaving plot margins unchanged.
const AXIS_TITLE_MIN_X = 12;

const numericAxisLabelMaxWidth = (canvasWidth: number): number =>
  Math.min(AXIS_LABEL_MAX_WIDTH, Math.max(12, (canvasWidth * 0.75 - 40) / 1.5));

/**
 * Keep host-supplied format strings from expanding the scene without bound.
 * The original value remains available to assistive consumers when the
 * visible SVG text needs an ellipsis.
 */
const boundedAxisLabel = (label: string, maxWidth = AXIS_LABEL_MAX_WIDTH): string => {
  if (label.length <= AXIS_LABEL_MAX_CHARS && textWidth(label) <= maxWidth) return label;
  let result = '';
  let characterCount = 0;
  for (const character of label) {
    if (characterCount >= AXIS_LABEL_MAX_CHARS) break;
    const candidate = `${result}${character}…`;
    if (textWidth(candidate) > maxWidth) break;
    result += character;
    characterCount += 1;
  }
  return result.length > 0 ? `${result}…` : '…';
};

const numericAxisLabel = (
  attrs: Record<string, string | number>,
  value: number,
  percent: boolean,
  axis: ChartAxisSpec | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  maxLabelWidth: number,
): ChartPrimitive => {
  const fullText = axisTickLabel(value, percent, axis, locale, formatter);
  const text = boundedAxisLabel(fullText, maxLabelWidth);
  return primitive(
    'text',
    {
      ...attrs,
      ...(text === fullText ? {} : { 'aria-label': fullText }),
    },
    'axis-label',
    undefined,
    undefined,
    text,
  );
};

const categoryAxisLabel = (
  attrs: Record<string, string | number>,
  label: string,
  position: ResolvedAxisLabelPosition,
  maxLabelWidth: number,
  categoryOriginalIndex: number,
  categoryId: string,
): ChartPrimitive => {
  const text = position === 'high' ? boundedAxisLabel(label, maxLabelWidth) : label;
  return primitive(
    'text',
    {
      ...attrs,
      ...(text === label ? {} : { 'aria-label': label }),
    },
    'category-label',
    undefined,
    categoryOriginalIndex,
    text,
    { pointId: categoryId },
  );
};

const numericAxisLabelVisibility = (
  ticks: readonly number[],
  positions: readonly number[],
  percent: boolean,
  axis: ChartAxisSpec | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  maxLabelWidth: number,
  orientation: 'horizontal' | 'vertical' = 'horizontal',
): readonly boolean[] => {
  const visible = Array.from({ length: ticks.length }, () => false);
  if (ticks.length === 0) return visible;
  if (ticks.length === 1) {
    visible[0] = true;
    return visible;
  }
  const sizes = ticks.map((tick) =>
    orientation === 'horizontal'
      ? textWidth(
          boundedAxisLabel(axisTickLabel(tick, percent, axis, locale, formatter), maxLabelWidth),
        )
      : 10,
  );
  const order = positions
    .map((_, index) => index)
    .sort((left, right) => (positions[left] ?? 0) - (positions[right] ?? 0));
  const gap = orientation === 'vertical' ? 3 : 4;
  const canSeparate = (leftIndex: number, rightIndex: number): boolean => {
    const leftPosition = positions[leftIndex] ?? 0;
    const rightPosition = positions[rightIndex] ?? 0;
    const leftWidth = sizes[leftIndex] ?? 0;
    const rightWidth = sizes[rightIndex] ?? 0;
    return rightPosition - rightWidth / 2 >= leftPosition + leftWidth / 2 + gap;
  };
  const first = order[0] ?? 0;
  const last = order[order.length - 1] ?? first;
  const kept = [first];
  visible[first] = true;
  for (let orderIndex = 1; orderIndex < order.length - 1; orderIndex += 1) {
    const index = order[orderIndex] ?? first;
    const previous = kept[kept.length - 1] ?? 0;
    if (canSeparate(previous, index)) {
      kept.push(index);
      visible[index] = true;
    }
  }
  while (kept.length > 1 && !canSeparate(kept[kept.length - 1] ?? first, last)) {
    const removed = kept.pop();
    if (removed !== undefined) visible[removed] = false;
  }
  if (canSeparate(kept[kept.length - 1] ?? first, last)) {
    visible[last] = true;
  }
  return visible;
};

const addAxis = (
  elements: ChartPrimitive[],
  plot: ChartScene['plot'],
  kind: ChartKind,
  xDomain: Domain | null,
  yDomain: Domain | null,
  xLabels: readonly string[],
  categoryOriginalIndices: readonly number[],
  categoryIds: readonly string[],
  spec: ChartSpec,
  percent: boolean,
  canvasHeight: number,
  maxLabelWidth: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  yAxisLabelMaxWidth = maxLabelWidth,
  xLogBase?: number,
  yLogBase?: number,
  axisTitle?: string,
): void => {
  const horizontalCategory = kind === 'bar';
  if (isCircularKind(kind)) return;

  const xAxis = spec.xAxis;
  const yAxis = spec.yAxis;
  const xVisible = axisIsVisible(xAxis);
  const yVisible = axisIsVisible(yAxis);
  const xGridlines = axisHasGridlines(xAxis);
  const yGridlines = axisHasGridlines(yAxis);
  const xLabelPosition = resolvedAxisLabelPosition(xAxis);
  const yLabelPosition = resolvedAxisLabelPosition(yAxis);
  const xScale = xDomain
    ? (value: number): number =>
        numericProjection(
          value,
          xDomain,
          plot.x,
          plot.x + plot.width,
          reversedAxis(xAxis),
          xLogBase,
        )
    : null;
  const yScale = yDomain
    ? (value: number): number =>
        numericProjection(
          value,
          yDomain,
          plot.y + plot.height,
          plot.y,
          reversedAxis(yAxis),
          yLogBase,
        )
    : null;
  const categoryReverse = horizontalCategory ? reversedAxis(yAxis) : reversedAxis(xAxis);
  const categoryLabelMaxWidth = (axis: 'x' | 'y'): number => {
    if (axis === 'y') return yAxisLabelMaxWidth;
    if (xLabelPosition !== 'high') return maxLabelWidth;
    const categoryCount = Math.max(1, xLabels.length);
    return Math.min(maxLabelWidth, Math.max(12, plot.width / categoryCount - 4));
  };

  if (horizontalCategory && xDomain && xScale) {
    const zeroX = numericAxisCrossing(xAxis, xDomain, xScale, xLogBase);
    const categoryY = categoryAxisCrossing(
      yAxis,
      xLabels.length,
      plot.y,
      plot.y + plot.height,
      categoryReverse,
      plot.y + plot.height,
    );
    if (xVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: plot.x,
            x2: plot.x + plot.width,
            y1: categoryY,
            y2: categoryY,
            stroke: 'currentColor',
            'data-chart-axis': 'x',
          },
          'axis',
        ),
      );
    }
    const xTicks = axisTicks(xDomain, xAxis, xLogBase);
    const xLabelVisibility = xVisible
      ? numericAxisLabelVisibility(
          xTicks,
          xTicks.map((tick) => xScale(tick)),
          percent,
          xAxis,
          locale,
          formatter,
          maxLabelWidth,
        )
      : [];
    for (let index = 0; index < xTicks.length; index += 1) {
      const tick = xTicks[index] ?? 0;
      const x = xScale(tick);
      if (xGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: x,
              x2: x,
              y1: plot.y,
              y2: plot.y + plot.height,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'x',
            },
            'axis-tick',
          ),
        );
      }
      if (xVisible && xLabelPosition !== 'none' && xLabelVisibility[index]) {
        const labelY =
          xLabelPosition === 'high'
            ? plot.y - 4
            : xLabelPosition === 'nextTo'
              ? categoryY + 12
              : plot.y + plot.height + 12;
        elements.push(
          numericAxisLabel(
            { x, y: labelY, 'text-anchor': 'middle', 'data-chart-axis': 'x' },
            tick,
            percent,
            xAxis,
            locale,
            formatter,
            maxLabelWidth,
          ),
        );
      }
    }
    if (yVisible) {
      if (yLabelPosition !== 'none') {
        for (let index = 0; index < xLabels.length; index += 1) {
          const labelX =
            yLabelPosition === 'high'
              ? plot.x + plot.width + 4
              : yLabelPosition === 'nextTo'
                ? zeroX - 4
                : plot.x - 4;
          elements.push(
            categoryAxisLabel(
              {
                x: labelX,
                y: yCategory(plot, xLabels.length, index, categoryReverse) + 3,
                'text-anchor': yLabelPosition === 'high' ? 'start' : 'end',
                'data-chart-axis': 'y',
              },
              xLabels[index] ?? '',
              yLabelPosition,
              categoryLabelMaxWidth('y'),
              categoryOriginalIndices[index] ?? index,
              categoryIds[index] ?? `category-${categoryOriginalIndices[index] ?? index}`,
            ),
          );
        }
      }
      elements.push(
        primitive(
          'line',
          {
            x1: zeroX,
            x2: zeroX,
            y1: plot.y,
            y2: plot.y + plot.height,
            stroke: 'currentColor',
            'data-chart-axis': 'y',
          },
          'axis',
        ),
      );
    }
  } else if (!horizontalCategory && yDomain && yScale && !xDomain) {
    const zeroY = numericAxisCrossing(yAxis, yDomain, yScale, yLogBase);
    const categoryX = categoryAxisCrossing(
      xAxis,
      xLabels.length,
      plot.x,
      plot.x + plot.width,
      categoryReverse,
      plot.x,
    );
    if (xVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: plot.x,
            x2: plot.x + plot.width,
            y1: zeroY,
            y2: zeroY,
            stroke: 'currentColor',
            'data-chart-axis': 'x',
          },
          'axis',
        ),
      );
    }
    const yTicks = axisTicks(yDomain, yAxis, yLogBase);
    const yLabelVisibility = yVisible
      ? numericAxisLabelVisibility(
          yTicks,
          yTicks.map((tick) => yScale(tick)),
          percent,
          yAxis,
          locale,
          formatter,
          yAxisLabelMaxWidth,
          'vertical',
        )
      : [];
    for (let index = 0; index < yTicks.length; index += 1) {
      const tick = yTicks[index] ?? 0;
      const y = yScale(tick);
      if (yGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: plot.x,
              x2: plot.x + plot.width,
              y1: y,
              y2: y,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'y',
            },
            'axis-tick',
          ),
        );
      }
      if (yVisible && yLabelPosition !== 'none' && yLabelVisibility[index]) {
        const labelX =
          yLabelPosition === 'high'
            ? plot.x + plot.width + 4
            : yLabelPosition === 'nextTo'
              ? categoryX - 2
              : plot.x - 2;
        elements.push(
          numericAxisLabel(
            {
              x: labelX,
              y: y + 3,
              'text-anchor': yLabelPosition === 'high' ? 'start' : 'end',
              'data-chart-axis': 'y',
            },
            tick,
            percent,
            yAxis,
            locale,
            formatter,
            yAxisLabelMaxWidth,
          ),
        );
      }
    }
    if (xVisible) {
      for (let index = 0; index < xLabels.length; index += 1) {
        if (xLabelPosition === 'none') continue;
        const labelY =
          xLabelPosition === 'high'
            ? plot.y - 4
            : xLabelPosition === 'nextTo'
              ? zeroY + 12
              : plot.y + plot.height + 12;
        elements.push(
          categoryAxisLabel(
            {
              x: xCategory(plot, xLabels.length, index, categoryReverse),
              y: labelY,
              'text-anchor': 'middle',
              'data-chart-axis': 'x',
            },
            xLabels[index] ?? '',
            xLabelPosition,
            categoryLabelMaxWidth('x'),
            categoryOriginalIndices[index] ?? index,
            categoryIds[index] ?? `category-${categoryOriginalIndices[index] ?? index}`,
          ),
        );
      }
    }
    if (yVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: categoryX,
            x2: categoryX,
            y1: plot.y,
            y2: plot.y + plot.height,
            stroke: 'currentColor',
            'data-chart-axis': 'y',
          },
          'axis',
        ),
      );
    }
  } else if (xDomain && xScale && yDomain && yScale) {
    const zeroY = numericAxisCrossing(yAxis, yDomain, yScale, yLogBase);
    const zeroX = numericAxisCrossing(xAxis, xDomain, xScale, xLogBase);
    if (xVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: plot.x,
            x2: plot.x + plot.width,
            y1: zeroY,
            y2: zeroY,
            stroke: 'currentColor',
            'data-chart-axis': 'x',
          },
          'axis',
        ),
      );
    }
    if (yVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: zeroX,
            x2: zeroX,
            y1: plot.y,
            y2: plot.y + plot.height,
            stroke: 'currentColor',
            'data-chart-axis': 'y',
          },
          'axis',
        ),
      );
    }
    const yTicks = axisTicks(yDomain, yAxis, yLogBase);
    const yLabelVisibility = yVisible
      ? numericAxisLabelVisibility(
          yTicks,
          yTicks.map((tick) => yScale(tick)),
          percent,
          yAxis,
          locale,
          formatter,
          yAxisLabelMaxWidth,
          'vertical',
        )
      : [];
    for (let index = 0; index < yTicks.length; index += 1) {
      const tick = yTicks[index] ?? 0;
      const y = yScale(tick);
      if (yGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: plot.x,
              x2: plot.x + plot.width,
              y1: y,
              y2: y,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'y',
            },
            'axis-tick',
          ),
        );
      }
      if (yVisible && yLabelPosition !== 'none' && yLabelVisibility[index]) {
        const labelX =
          yLabelPosition === 'high'
            ? plot.x + plot.width + 4
            : yLabelPosition === 'nextTo'
              ? zeroX - 2
              : plot.x - 2;
        elements.push(
          numericAxisLabel(
            {
              x: labelX,
              y: y + 3,
              'text-anchor': yLabelPosition === 'high' ? 'start' : 'end',
              'data-chart-axis': 'y',
            },
            tick,
            percent,
            yAxis,
            locale,
            formatter,
            yAxisLabelMaxWidth,
          ),
        );
      }
    }
    const xTicks = axisTicks(xDomain, xAxis, xLogBase);
    const xLabelVisibility = xVisible
      ? numericAxisLabelVisibility(
          xTicks,
          xTicks.map((tick) => xScale(tick)),
          false,
          xAxis,
          locale,
          formatter,
          maxLabelWidth,
        )
      : [];
    for (let index = 0; index < xTicks.length; index += 1) {
      const tick = xTicks[index] ?? 0;
      const x = xScale(tick);
      if (xGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: x,
              x2: x,
              y1: plot.y,
              y2: plot.y + plot.height,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'x',
            },
            'axis-tick',
          ),
        );
      }
      if (xVisible && xLabelPosition !== 'none' && xLabelVisibility[index]) {
        const labelY =
          xLabelPosition === 'high'
            ? plot.y - 4
            : xLabelPosition === 'nextTo'
              ? zeroY + 12
              : plot.y + plot.height + 12;
        elements.push(
          numericAxisLabel(
            { x, y: labelY, 'text-anchor': 'middle', 'data-chart-axis': 'x' },
            tick,
            false,
            xAxis,
            locale,
            formatter,
            maxLabelWidth,
          ),
        );
      }
    }
  }

  const xTitle = xVisible ? axisTitleText(xAxis?.title, axisTitle) : undefined;
  const yTitle = yVisible ? axisTitleText(yAxis?.title, axisTitle) : undefined;
  if (xTitle) {
    elements.push(
      primitive(
        'text',
        {
          x: plot.x + plot.width / 2,
          y: Math.min(canvasHeight - 2, plot.y + plot.height + 25),
          'text-anchor': 'middle',
          'data-chart-axis': 'x',
        },
        'axis-title',
        undefined,
        undefined,
        xTitle,
      ),
    );
  }
  if (yTitle) {
    const yLabelWidth =
      yLabelPosition === 'high'
        ? 12
        : horizontalCategory
          ? xLabels.reduce((current, label) => Math.max(current, textWidth(label)), 12)
          : yDomain
            ? axisTicks(yDomain, yAxis, yLogBase).reduce(
                (current, tick) =>
                  Math.max(
                    current,
                    textWidth(
                      boundedAxisLabel(
                        axisTickLabel(tick, percent, yAxis, locale, formatter),
                        yAxisLabelMaxWidth,
                      ),
                    ),
                  ),
                12,
              )
            : 12;
    const x = Math.max(AXIS_TITLE_MIN_X, plot.x - Math.max(28, yLabelWidth + 16));
    const y = plot.y + plot.height / 2;
    elements.push(
      primitive(
        'text',
        {
          x,
          y,
          transform: `rotate(-90 ${x} ${y})`,
          'text-anchor': 'middle',
          'data-chart-axis': 'y',
        },
        'axis-title',
        undefined,
        undefined,
        yTitle,
      ),
    );
  }
};

/**
 * Sum ordinary stack totals in the finite spreadsheet number range. If the
 * mathematical sum exceeds that range, the representable axis is clipped at
 * ±MAX_VALUE; percent stacks use stableStackTotal below so their ratios do
 * not depend on that clipping.
 */
const effectiveStackTotals = (
  data: CategoryChartData,
): { positive: number[]; negative: number[] } => {
  const positive = Array.from({ length: data.categories.length }, () => 0);
  const negative = Array.from({ length: data.categories.length }, () => 0);
  for (const series of data.series) {
    for (let index = 0; index < data.categories.length; index += 1) {
      const value = valueFor(series.values[index] ?? null);
      if (value === null) continue;
      if (value >= 0) {
        positive[index] = saturatedAdd(positive[index] ?? 0, value);
      } else {
        negative[index] = saturatedAdd(negative[index] ?? 0, value);
      }
    }
  }
  return { positive, negative };
};

interface StableStackTotal {
  readonly scale: number;
  readonly total: number;
}

const stableStackTotal = (
  data: CategoryChartData,
  pointIndex: number,
  positive: boolean,
): StableStackTotal => {
  let scale = 0;
  for (const series of data.series) {
    const value = valueFor(series.values[pointIndex] ?? null);
    if (value === null || value >= 0 !== positive) continue;
    scale = Math.max(scale, Math.abs(value));
  }
  if (!(scale > 0)) return { scale: 1, total: 0 };
  let total = 0;
  for (const series of data.series) {
    const value = valueFor(series.values[pointIndex] ?? null);
    if (value === null || value >= 0 !== positive) continue;
    total = saturatedAdd(total, Math.abs(value) / scale);
  }
  return { scale, total };
};

interface StackSegment {
  readonly start: number;
  readonly end: number;
  readonly value: number;
}

const segmentsFor = (
  data: CategoryChartData,
  grouping: ChartGrouping,
): readonly (readonly (StackSegment | null)[])[] => {
  const stableTotals =
    grouping === 'percent-stacked'
      ? data.categories.map((_category, pointIndex) => ({
          positive: stableStackTotal(data, pointIndex, true),
          negative: stableStackTotal(data, pointIndex, false),
        }))
      : [];
  const result: (StackSegment | null)[][] = data.series.map(() =>
    Array.from({ length: data.categories.length }, () => null),
  );
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    const target = result[seriesIndex];
    if (!series || !target) continue;
    for (let pointIndex = 0; pointIndex < data.categories.length; pointIndex += 1) {
      const raw = series.values[pointIndex] ?? null;
      const value = valueFor(raw);
      if (value === null) continue;
      if (grouping === 'clustered') {
        target[pointIndex] = { start: 0, end: value, value };
        continue;
      }
      const positive = value >= 0;
      const stableTotal = (positive
        ? stableTotals[pointIndex]?.positive
        : stableTotals[pointIndex]?.negative) ?? {
        scale: 1,
        total: 0,
      };
      // Keep positive and negative accumulators separate by deriving them from
      // preceding segments. This avoids mutating the input data model.
      let start = 0;
      for (let prior = 0; prior < seriesIndex; prior += 1) {
        const priorValue = data.series[prior]?.values[pointIndex] ?? null;
        const normalized = valueFor(priorValue);
        if (normalized === null || normalized >= 0 !== positive) continue;
        start = saturatedAdd(start, normalized);
      }
      let rendered = value;
      if (grouping === 'percent-stacked') {
        rendered = stableTotal.total === 0 ? 0 : value / stableTotal.scale / stableTotal.total;
        start = 0;
        for (let prior = 0; prior < seriesIndex; prior += 1) {
          const priorValue = valueFor(data.series[prior]?.values[pointIndex] ?? null);
          if (priorValue === null || priorValue >= 0 !== positive) continue;
          start = saturatedAdd(
            start,
            stableTotal.total === 0 ? 0 : priorValue / stableTotal.scale / stableTotal.total,
          );
        }
      }
      target[pointIndex] = { start, end: saturatedAdd(start, rendered), value: rendered };
    }
  }
  return result;
};

interface BarGeometry {
  readonly size: number;
  readonly seriesStep: number;
  readonly groupOffset: number;
}

/** Resolve Excel's gapWidth/overlap equations for one category slot. */
const barGeometry = (
  slot: number,
  seriesCount: number,
  grouping: ChartGrouping,
  spec: ChartSpec,
): BarGeometry => {
  const gapWidth = boundedOption(spec.gapWidth, 0, 500, 150);
  const overlap = boundedOption(spec.overlap, -100, 100, 0);
  const count = Math.max(1, seriesCount);
  if (grouping !== 'clustered') {
    const size = slot / (1 + gapWidth / 100);
    return {
      size: finite(size, 0),
      seriesStep: 0,
      groupOffset: finite((slot - size) / 2, 0),
    };
  }
  const seriesStepFactor = 1 - overlap / 100;
  const size = slot / (1 + (count - 1) * seriesStepFactor + gapWidth / 100);
  const cluster = size + (count - 1) * size * seriesStepFactor;
  return {
    size: finite(size, 0),
    seriesStep: finite(size * seriesStepFactor, 0),
    groupOffset: finite((slot - cluster) / 2, 0),
  };
};

const addCategoryMarks = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  yDomain: Domain,
  palette: readonly string[],
  numericLogBase?: number,
): void => {
  const data = view.data;
  const grouping = spec.grouping ?? 'clustered';
  const segments = segmentsFor(data, grouping);
  const count = Math.max(1, data.categories.length);
  const isBar = spec.kind === 'bar';
  const slot = (isBar ? plot.height : plot.width) / count;
  const numericReverse = isBar ? reversedAxis(spec.xAxis) : reversedAxis(spec.yAxis);
  const categoryReverse = isBar ? reversedAxis(spec.yAxis) : reversedAxis(spec.xAxis);
  const numericScale = (value: number): number => {
    if (isBar) {
      return numericProjection(
        value,
        yDomain,
        plot.x,
        plot.x + plot.width,
        numericReverse,
        numericLogBase,
      );
    }
    return numericProjection(
      value,
      yDomain,
      plot.y + plot.height,
      plot.y,
      numericReverse,
      numericLogBase,
    );
  };
  if (spec.kind === 'column' || spec.kind === 'bar') {
    const seriesCount = Math.max(1, data.series.length);
    const geometry = barGeometry(slot, seriesCount, grouping, spec);
    for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
      const series = data.series[seriesIndex];
      if (!series) continue;
      const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
      for (let pointIndex = 0; pointIndex < data.categories.length; pointIndex += 1) {
        const segment = segments[seriesIndex]?.[pointIndex] ?? null;
        if (!segment) continue;
        if (
          numericLogBase !== undefined &&
          (!(segment.value > 0) || !Number.isFinite(segment.value))
        )
          continue;
        const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
        const pointId = categoryIdAt(view, pointIndex);
        const style = resolveChartStyle(
          spec,
          palette,
          series.id,
          pointId,
          originalSeriesIndex,
          originalPointIndex,
        );
        const valueStart = numericScale(
          numericLogBase !== undefined && !(segment.start > 0) ? yDomain.min : segment.start,
        );
        const valueEnd = numericScale(segment.end);
        const categoryPosition = categoryIndex(count, pointIndex, categoryReverse);
        if (isBar) {
          const x = Math.min(valueStart, valueEnd);
          elements.push(
            primitive(
              'rect',
              styleAttrs(
                {
                  x,
                  y:
                    plot.y +
                    categoryPosition * slot +
                    geometry.groupOffset +
                    seriesIndex * geometry.seriesStep,
                  width: Math.abs(valueEnd - valueStart),
                  height: geometry.size,
                },
                style,
              ),
              'bar',
              originalSeriesIndex,
              originalPointIndex,
              undefined,
              { seriesId: series.id, pointId },
            ),
          );
        } else {
          const x =
            plot.x +
            categoryPosition * slot +
            geometry.groupOffset +
            seriesIndex * geometry.seriesStep;
          elements.push(
            primitive(
              'rect',
              styleAttrs(
                {
                  x,
                  y: Math.min(valueStart, valueEnd),
                  width: geometry.size,
                  height: Math.abs(valueEnd - valueStart),
                },
                style,
              ),
              'bar',
              originalSeriesIndex,
              originalPointIndex,
              undefined,
              { seriesId: series.id, pointId },
            ),
          );
        }
      }
    }
    return;
  }

  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    const points = Array.from({ length: data.categories.length }, (_value, pointIndex) => {
      const segment = segments[seriesIndex]?.[pointIndex] ?? null;
      const plottableSegment =
        segment !== null &&
        (numericLogBase === undefined || (segment.value > 0 && Number.isFinite(segment.value)));
      return {
        pointIndex,
        x: xCategory(plot, data.categories.length, pointIndex, categoryReverse),
        y: plottableSegment ? numericScale(segment.end) : null,
        segment: plottableSegment ? segment : null,
      };
    });
    const runs: { index: number; value: number; base: number }[][] = [];
    let run: { index: number; value: number; base: number }[] = [];
    for (const point of points) {
      if (point.y === null || !point.segment) {
        const raw = data.series[seriesIndex]?.values[point.pointIndex] ?? null;
        const logBreak =
          numericLogBase !== undefined && raw !== null && (!Number.isFinite(raw) || raw <= 0);
        if (logBreak) {
          if (run.length > 0) runs.push(run);
          run = [];
          continue;
        }
        if (series.states?.[point.pointIndex] === 'connect') continue;
        if (run.length > 0) runs.push(run);
        run = [];
      } else {
        run.push({
          index: point.pointIndex,
          value: point.y,
          base: numericScale(
            numericLogBase !== undefined && !(point.segment.start > 0)
              ? yDomain.min
              : point.segment.start,
          ),
        });
      }
    }
    if (run.length > 0) runs.push(run);
    for (const lineRun of runs) {
      const mapped = lineRun
        .map(
          (point) =>
            `${xCategory(plot, data.categories.length, point.index, categoryReverse)},${point.value}`,
        )
        .join(' ');
      if (spec.kind === 'area' && lineRun.length > 0) {
        const lower = lineRun
          .slice()
          .reverse()
          .map(
            (point) =>
              `${xCategory(plot, data.categories.length, point.index, categoryReverse)},${point.base}`,
          )
          .join(' ');
        const style = resolveChartStyle(
          spec,
          palette,
          series.id,
          undefined,
          originalSeriesIndex,
          undefined,
        );
        elements.push(
          primitive(
            'polygon',
            styleAttrs(
              {
                points: `${mapped} ${lower}`,
              },
              style,
              0.2,
            ),
            'area',
            originalSeriesIndex,
            undefined,
            undefined,
            { seriesId: series.id },
          ),
        );
      }
      if (lineRun.length >= 2) {
        const style = resolveChartStyle(
          spec,
          palette,
          series.id,
          undefined,
          originalSeriesIndex,
          undefined,
        );
        const opacity = lineOpacity(style.line);
        const width = lineWidth(style.line, 2) ?? 2;
        const dash = dashArray(style.line?.dash);
        const smooth = spec.kind === 'line' && smoothed(spec, series.id);
        const coordinates = lineRun.map(
          (point) =>
            [
              xCategory(plot, data.categories.length, point.index, categoryReverse),
              point.value,
            ] as const,
        );
        elements.push(
          primitive(
            smooth ? 'path' : 'polyline',
            {
              ...(smooth ? { d: smoothPath(coordinates) } : { points: mapped }),
              fill: 'none',
              stroke: style.line?.color ?? style.color,
              'stroke-width': width,
              ...(opacity === undefined ? {} : { 'stroke-opacity': opacity }),
              ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
            },
            'line',
            originalSeriesIndex,
            undefined,
            undefined,
            { seriesId: series.id },
          ),
        );
      }
    }
    for (const point of points) {
      if (point.y === null) continue;
      const originalPointIndex = view.originalCategoryIndices[point.pointIndex] ?? point.pointIndex;
      const pointId = categoryIdAt(view, point.pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      if (!markerVisible(style.marker)) continue;
      const marker = markerAttrs(point.x, point.y, style);
      elements.push(
        primitive(
          marker.tag,
          marker.attrs,
          'point',
          originalSeriesIndex,
          originalPointIndex,
          undefined,
          { seriesId: series.id, pointId },
        ),
      );
    }
  }
};

const addScatterMarks = (
  elements: ChartPrimitive[],
  view: XYChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  xDomain: Domain,
  yDomain: Domain,
  palette: readonly string[],
  xLogBase?: number,
  yLogBase?: number,
): void => {
  const data = view.data;
  const xScale = (value: number): number =>
    numericProjection(
      value,
      xDomain,
      plot.x,
      plot.x + plot.width,
      reversedAxis(spec.xAxis),
      xLogBase,
    );
  const yScale = (value: number): number =>
    numericProjection(
      value,
      yDomain,
      plot.y + plot.height,
      plot.y,
      reversedAxis(spec.yAxis),
      yLogBase,
    );
  const inDomain = (value: number, domain: Domain): boolean =>
    Number.isFinite(value) && value >= domain.min && value <= domain.max;
  const logEligible = (value: number, logBase: number | undefined): boolean =>
    logBase === undefined || value > 0;
  const lineAxis = (domain: Domain, logBase: number | undefined) => {
    if (logBase === undefined) return { domain, transform: (value: number): number => value };
    const domainMin = Math.log(domain.min);
    const domainMax = Math.log(domain.max);
    if (!(domainMax > domainMin) || !Number.isFinite(domainMin) || !Number.isFinite(domainMax)) {
      return { domain, transform: (value: number): number => value };
    }
    return {
      domain: { min: domainMin, max: domainMax },
      transform: (value: number): number => Math.log(value),
    };
  };
  const lineX = lineAxis(xDomain, xLogBase);
  const lineY = lineAxis(yDomain, yLogBase);
  const lineDomains = { x: lineX.domain, y: lineY.domain };
  const scatterStyle = spec.scatterStyle ?? 'marker';
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    const lineSamples: ScatterLineSample[] = [];
    const markerPoints: Array<{
      readonly originalPointIndex: number;
      readonly pointId: string;
      readonly x: number;
      readonly y: number;
    }> = [];
    for (let pointIndex = 0; pointIndex < series.points.length; pointIndex += 1) {
      const point = series.points[pointIndex];
      if (!point) continue;
      const originalPointIndex = view.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex;
      const pointId = pointIdAt(view, seriesIndex, pointIndex);
      if (point.state === 'connect') continue;
      const nonFinite =
        (point.x !== null && !Number.isFinite(point.x)) ||
        (point.y !== null && !Number.isFinite(point.y));
      const nonPositiveLog =
        (point.x !== null && Number.isFinite(point.x) && !logEligible(point.x, xLogBase)) ||
        (point.y !== null && Number.isFinite(point.y) && !logEligible(point.y, yLogBase));
      if (nonFinite || nonPositiveLog) {
        // Non-finite values are a hard run break, even next to a connected gap.
        lineSamples.push({ kind: 'break', sourceIndex: originalPointIndex });
        continue;
      }
      if (point.x === null || point.y === null) {
        lineSamples.push({ kind: 'blank', sourceIndex: originalPointIndex });
        continue;
      }
      lineSamples.push({
        kind: 'point',
        x: lineX.transform(point.x),
        y: lineY.transform(point.y),
        sourceIndex: originalPointIndex,
      });
      if (
        inDomain(point.x, xDomain) &&
        inDomain(point.y, yDomain) &&
        logEligible(point.x, xLogBase) &&
        logEligible(point.y, yLogBase)
      ) {
        markerPoints.push({ originalPointIndex, pointId, x: point.x, y: point.y });
      }
    }

    if (scatterStyle !== 'marker') {
      const lineStyle = resolveChartStyle(
        spec,
        palette,
        series.id,
        undefined,
        originalSeriesIndex,
        undefined,
      );
      const width = lineWidth(lineStyle.line, 2) ?? 2;
      // A zero-width series line is an explicit request to suppress the
      // connector while leaving point markers independently controllable.
      if (width > 0) {
        const opacity = lineOpacity(lineStyle.line);
        const dash = dashArray(lineStyle.line?.dash);
        const stroke = {
          fill: 'none',
          stroke: lineStyle.line?.color ?? lineStyle.color,
          'stroke-width': width,
          ...(opacity === undefined ? {} : { 'stroke-opacity': opacity }),
          ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
        };
        const segments = scatterLineSegments(lineSamples, plot, false, lineDomains, {
          reverseX: reversedAxis(spec.xAxis),
          reverseY: reversedAxis(spec.yAxis),
        });
        if (smoothed(spec, series.id)) {
          // A smoothed line runs through each chain of joined segments.
          const runs: [number, number][][] = [];
          for (const segment of segments) {
            const run = runs.at(-1);
            const end = run?.at(-1);
            if (run && end && end[0] === segment.x1 && end[1] === segment.y1) {
              run.push([segment.x2, segment.y2]);
            } else {
              runs.push([
                [segment.x1, segment.y1],
                [segment.x2, segment.y2],
              ]);
            }
          }
          for (const run of runs) {
            elements.push(
              primitive(
                'path',
                { d: smoothPath(run), ...stroke },
                'line',
                originalSeriesIndex,
                undefined,
                undefined,
                {
                  seriesId: series.id,
                },
              ),
            );
          }
        } else {
          for (const segment of segments) {
            elements.push(
              primitive(
                'line',
                { x1: segment.x1, y1: segment.y1, x2: segment.x2, y2: segment.y2, ...stroke },
                'line',
                originalSeriesIndex,
                undefined,
                undefined,
                { seriesId: series.id },
              ),
            );
          }
        }
      }
    }

    for (const point of markerPoints) {
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        point.pointId,
        originalSeriesIndex,
        point.originalPointIndex,
      );
      if (!markerVisible(style.marker)) continue;
      const marker = markerAttrs(xScale(point.x), yScale(point.y), style);
      elements.push(
        primitive(
          marker.tag,
          marker.attrs,
          'point',
          originalSeriesIndex,
          point.originalPointIndex,
          undefined,
          { seriesId: series.id, pointId: point.pointId },
        ),
      );
    }
  }
};

const addPieMarks = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  plot: ChartScene['plot'],
  palette: readonly string[],
  spec: ChartSpec,
): void => {
  const data = view.data;
  const first = data.series[0];
  if (!first) return;
  const magnitudes = first.values.map((value) =>
    value !== null && Number.isFinite(value) && value !== 0 ? Math.abs(value) : 0,
  );
  const scale = magnitudes.reduce((maximum, value) => Math.max(maximum, value), 0);
  if (!(scale > 0)) return;
  const normalized = magnitudes.map((value) => value / scale);
  const total = normalized.reduce((sum, value) => sum + value, 0);
  if (!(total > 0)) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const radius = Math.max(0, Math.min(plot.width, plot.height) / 2 - 2);
  const firstSliceAngle = boundedOption(spec.firstSliceAngle, 0, 360, 0);
  let start = -Math.PI / 2 + (firstSliceAngle * Math.PI) / 180;
  for (let pointIndex = 0; pointIndex < normalized.length; pointIndex += 1) {
    const value = normalized[pointIndex] ?? 0;
    if (!(value > 0)) continue;
    const angle = (value / total) * Math.PI * 2;
    const end = start + angle;
    const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
    const pointId = categoryIdAt(view, pointIndex);
    const style = resolveChartStyle(
      spec,
      palette,
      first.id,
      pointId,
      view.originalSeriesIndices[0] ?? 0,
      originalPointIndex,
    );
    if (angle >= Math.PI * 2 - 1e-9) {
      elements.push(
        primitive(
          'circle',
          styleAttrs({ cx, cy, r: radius }, style),
          'slice',
          view.originalSeriesIndices[0] ?? 0,
          originalPointIndex,
          undefined,
          { seriesId: first.id, pointId },
        ),
      );
    } else {
      const startX = cx + Math.cos(start) * radius;
      const startY = cy + Math.sin(start) * radius;
      const endX = cx + Math.cos(end) * radius;
      const endY = cy + Math.sin(end) * radius;
      elements.push(
        primitive(
          'path',
          {
            d: `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${angle > Math.PI ? 1 : 0} 1 ${endX} ${endY} Z`,
            ...styleAttrs(
              {
                stroke: 'Canvas',
                'stroke-width': 1,
              },
              style,
            ),
          },
          'slice',
          view.originalSeriesIndices[0] ?? 0,
          originalPointIndex,
          undefined,
          { seriesId: first.id, pointId },
        ),
      );
    }
    start = end;
  }
};

interface NormalizedMagnitudes {
  readonly values: readonly number[];
  readonly total: number;
}

/** Normalize one circular ring without overflowing on spreadsheet extremes. */
const normalizedMagnitudes = (values: readonly (number | null)[]): NormalizedMagnitudes | null => {
  const magnitudes = values.map((value) =>
    value !== null && Number.isFinite(value) && value !== 0 ? Math.abs(value) : 0,
  );
  const scale = magnitudes.reduce((maximum, value) => Math.max(maximum, value), 0);
  if (!(scale > 0) || !Number.isFinite(scale)) return null;
  const normalized = magnitudes.map((value) => value / scale);
  let total = 0;
  for (const value of normalized) {
    total += value;
    if (!Number.isFinite(total)) return null;
  }
  return total > 0 ? { values: normalized, total } : null;
};

interface CircularPoint {
  readonly x: number;
  readonly y: number;
}

const circularPoint = (cx: number, cy: number, radius: number, angle: number): CircularPoint => ({
  x: cx + Math.cos(angle) * radius,
  y: cy + Math.sin(angle) * radius,
});

/** Build an annular SVG path, including two arcs for each side of a full ring. */
const annularPath = (
  cx: number,
  cy: number,
  innerRadius: number,
  outerRadius: number,
  start: number,
  angle: number,
): string => {
  const fullCircle = angle >= Math.PI * 2 - 1e-9;
  const outerStart = circularPoint(cx, cy, outerRadius, start);
  const innerStart = circularPoint(cx, cy, innerRadius, start);
  if (fullCircle) {
    const outerMid = circularPoint(cx, cy, outerRadius, start + Math.PI);
    const innerMid = circularPoint(cx, cy, innerRadius, start + Math.PI);
    return [
      `M ${outerStart.x} ${outerStart.y}`,
      `A ${outerRadius} ${outerRadius} 0 1 1 ${outerMid.x} ${outerMid.y}`,
      `A ${outerRadius} ${outerRadius} 0 1 1 ${outerStart.x} ${outerStart.y}`,
      `M ${innerStart.x} ${innerStart.y}`,
      `A ${innerRadius} ${innerRadius} 0 1 0 ${innerMid.x} ${innerMid.y}`,
      `A ${innerRadius} ${innerRadius} 0 1 0 ${innerStart.x} ${innerStart.y}`,
      'Z',
    ].join(' ');
  }
  const end = start + angle;
  const outerEnd = circularPoint(cx, cy, outerRadius, end);
  const innerEnd = circularPoint(cx, cy, innerRadius, end);
  const largeArc = angle > Math.PI ? 1 : 0;
  return [
    `M ${outerStart.x} ${outerStart.y}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${outerEnd.x} ${outerEnd.y}`,
    `L ${innerEnd.x} ${innerEnd.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${innerStart.x} ${innerStart.y}`,
    'Z',
  ].join(' ');
};

const addDoughnutMarks = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  plot: ChartScene['plot'],
  palette: readonly string[],
  spec: ChartSpec,
): void => {
  const data = view.data;
  if (data.series.length === 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const outerRadius = Math.max(0, Math.min(plot.width, plot.height) / 2 - 2);
  const holePercent = boundedOption(spec.holeSize, 10, 90, 50);
  const holeRadius = (outerRadius * holePercent) / 100;
  const ringWidth = (outerRadius - holeRadius) / data.series.length;
  const firstSliceAngle = boundedOption(spec.firstSliceAngle, 0, 360, 0);
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const normalized = normalizedMagnitudes(series.values);
    if (!normalized) continue;
    const innerRadius = holeRadius + ringWidth * seriesIndex;
    const outerRingRadius = innerRadius + ringWidth;
    let start = -Math.PI / 2 + (firstSliceAngle * Math.PI) / 180;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < normalized.values.length; pointIndex += 1) {
      const value = normalized.values[pointIndex] ?? 0;
      if (!(value > 0) || !Number.isFinite(value)) continue;
      const angle = (value / normalized.total) * Math.PI * 2;
      if (!(angle > 0) || !Number.isFinite(angle)) continue;
      const end = start + angle;
      const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
      const pointId = categoryIdAt(view, pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      elements.push(
        primitive(
          'path',
          styleAttrs(
            {
              d: annularPath(cx, cy, innerRadius, outerRingRadius, start, angle),
              'fill-rule': 'evenodd',
              stroke: 'Canvas',
              'stroke-width': 1,
            },
            style,
          ),
          'slice',
          originalSeriesIndex,
          originalPointIndex,
          undefined,
          { seriesId: series.id, pointId },
        ),
      );
      start = end;
    }
  }
};

const addLegend = (
  elements: ChartPrimitive[],
  view: PreparedChartData,
  spec: ChartSpec,
  width: number,
  height: number,
  plot: ChartScene['plot'],
  palette: readonly string[],
  rightLabelGutter = 0,
): void => {
  const data = view.data;
  const entries: {
    label: string;
    color: string;
    seriesIndex?: number;
    pointIndex?: number;
    seriesId?: string;
    pointId?: string;
  }[] = [];
  if (data.domain === 'category' && isCircularKind(spec.kind)) {
    const categoryView = view as CategoryChartView;
    if (spec.kind === 'doughnut' && data.series.length === 0) return;
    for (let index = 0; index < data.categories.length; index += 1) {
      const originalPointIndex = categoryView.originalCategoryIndices[index] ?? index;
      const pointId = categoryIdAt(categoryView, index);
      const series = data.series[0];
      const style = series
        ? resolveChartStyle(
            spec,
            palette,
            series.id,
            pointId,
            categoryView.originalSeriesIndices[0] ?? 0,
            originalPointIndex,
          )
        : undefined;
      entries.push({
        label: data.categories[index] ?? `${index + 1}`,
        color: style ? fillColor(style.fill, style.color) : colorFor(palette, originalPointIndex),
        seriesIndex: categoryView.originalSeriesIndices[0] ?? 0,
        pointIndex: originalPointIndex,
        seriesId: series?.id,
        pointId,
      });
    }
  } else if (data.domain === 'category') {
    const categoryView = view as CategoryChartView;
    data.series.forEach((series, index) => {
      const originalSeriesIndex = categoryView.originalSeriesIndices[index] ?? index;
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        undefined,
        originalSeriesIndex,
        undefined,
      );
      entries.push({
        label: series.name ?? `${originalSeriesIndex + 1}`,
        color: keyColor(spec, style),
        seriesIndex: originalSeriesIndex,
        seriesId: series.id,
      });
    });
  } else {
    data.series.forEach((series, index) => {
      const originalSeriesIndex = view.originalSeriesIndices[index] ?? index;
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        undefined,
        originalSeriesIndex,
        undefined,
      );
      entries.push({
        label: series.name ?? `${originalSeriesIndex + 1}`,
        color: keyColor(spec, style),
        seriesIndex: originalSeriesIndex,
        seriesId: series.id,
      });
    });
  }
  if (entries.length === 0) return;
  const position = spec.legend?.position ?? 'bottom';
  let x = plot.x;
  let y = 12;
  let horizontal = false;
  if (position === 'right') {
    x = plot.x + plot.width + 8 + rightLabelGutter;
    y = plot.y + 8;
  } else if (position === 'left') {
    x = 4;
    y = plot.y + 8;
  } else if (position === 'bottom') {
    x = 4;
    y = height - 5;
    horizontal = true;
  } else {
    x = 4;
    y = spec.title ? TITLE_RESERVE + 5 : 11;
    horizontal = true;
  }
  const availableWidth = Math.max(
    12,
    position === 'right' ? width - x - 4 : position === 'left' ? x + plot.x - 4 : width - x - 4,
  );
  const availableHeight = Math.max(14, position === 'bottom' ? height - 4 : height - y - 4);
  const preferredWidth = entries.reduce(
    (current, entry) => Math.max(current, textWidth(entry.label) + 20),
    56,
  );
  const columnWidth = Math.max(20, Math.min(preferredWidth, availableWidth));
  const verticalRows = Math.max(1, Math.floor(availableHeight / 14));
  const capacity = horizontal
    ? Math.max(1, Math.floor(availableWidth / 56))
    : Math.max(1, Math.floor(availableWidth / columnWidth)) * verticalRows;
  const visibleEntries =
    entries.length > capacity
      ? [
          ...entries.slice(0, Math.max(1, capacity - 1)),
          {
            label: `+${entries.length - Math.max(1, capacity - 1)}`,
            color: 'currentColor',
            seriesIndex: undefined,
            pointIndex: undefined,
            seriesId: undefined,
            pointId: undefined,
          },
        ]
      : entries;
  const items = visibleEntries.map((entry) => {
    const itemMaxWidth = horizontal
      ? Math.max(12, availableWidth / Math.max(1, visibleEntries.length) - 20)
      : Math.max(12, columnWidth - 11);
    const displayLabel = truncateLabel(entry.label, itemMaxWidth);
    return { entry, displayLabel, width: 20 + Math.min(itemMaxWidth, textWidth(displayLabel)) };
  });
  // A horizontal legend wraps into rows, each centred on the chart as Excel does.
  const rows: (typeof items)[] = [];
  if (horizontal) {
    let rowWidth = 0;
    for (const item of items) {
      const current = rows[rows.length - 1];
      if (current && rowWidth + item.width <= width - 8) {
        current.push(item);
        rowWidth += item.width;
      } else {
        rows.push([item]);
        rowWidth = item.width;
      }
    }
  }
  const placed = horizontal
    ? rows.flatMap((rowItems, row) => {
        // The trailing gap after the last label is not part of the row.
        const rowWidth = rowItems.reduce((sum, item) => sum + item.width, 0) - 9;
        let cursor = Math.max(4, (width - rowWidth) / 2);
        // A bottom legend grows upwards so its first row stays nearest the plot.
        const itemY = position === 'bottom' ? y - (rows.length - 1 - row) * 14 : y + row * 14;
        return rowItems.map((item) => {
          const itemX = cursor;
          cursor += item.width;
          return { ...item, itemX, itemY };
        });
      })
    : items.map((item, index) => ({
        ...item,
        itemX: x + Math.floor(index / verticalRows) * columnWidth,
        itemY: y + (index % verticalRows) * 14,
      }));
  for (const { entry, displayLabel, itemX, itemY } of placed) {
    elements.push(
      primitive(
        'rect',
        { x: itemX, y: itemY - 8, width: 8, height: 8, fill: entry.color },
        'legend',
        entry.seriesIndex,
        entry.pointIndex,
        undefined,
        { seriesId: entry.seriesId, pointId: entry.pointId },
      ),
      primitive(
        'text',
        { x: itemX + 11, y: itemY - 1, 'aria-label': entry.label },
        'legend',
        entry.seriesIndex,
        entry.pointIndex,
        displayLabel,
        { seriesId: entry.seriesId, pointId: entry.pointId },
      ),
    );
  }
};

interface LabelAnchor {
  readonly x: number;
  readonly y: number;
  readonly textAnchor?: 'start' | 'middle' | 'end';
}

const labelFontSize = (value: number | undefined): number =>
  value !== undefined && Number.isFinite(value) ? clamp(value, 6, 72) : 10;

const labelParts = (
  spec: NonNullable<ChartSpec['dataLabels']>,
  seriesName: string,
  categoryName: string | undefined,
  value: number | undefined,
  percentage: number | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): string => {
  const hasConfiguredLabelFlag = [
    spec.showSeriesName,
    spec.showCategoryName,
    spec.showValue,
    spec.showPercentage,
    spec.showLegendKey,
  ].some((flag) => flag !== undefined);
  const showValue = spec.showValue ?? !hasConfiguredLabelFlag;
  const parts: string[] = [];
  if (spec.showSeriesName) parts.push(seriesName);
  if (spec.showCategoryName && categoryName !== undefined) parts.push(categoryName);
  if (showValue && value !== undefined) {
    const formatted = formatChartValue(value, spec.numberFormat, locale, formatter);
    if (formatted !== null) parts.push(formatted);
  }
  if (spec.showPercentage && percentage !== undefined) {
    const percentageCode = spec.numberFormat?.includes('%')
      ? spec.numberFormat
      : spec.numberFormat === undefined
        ? '0%'
        : `${spec.numberFormat}%`;
    const formatted = formatChartValue(percentage, percentageCode, locale, formatter);
    if (formatted !== null) parts.push(formatted);
  }
  return parts.join(spec.separator ?? ', ');
};

const addDataLabel = (
  elements: ChartPrimitive[],
  fullText: string,
  anchor: LabelAnchor,
  spec: NonNullable<ChartSpec['dataLabels']>,
  width: number,
  height: number,
  seriesIndex: number,
  pointIndex: number,
  seriesId: string,
  pointId: string | undefined,
  color: string,
): void => {
  if (fullText.length === 0 && !spec.showLegendKey) return;
  const horizontalPadding = 1;
  const verticalPadding = 1;
  const glyphAscentRatio = 0.99;
  const glyphDescentRatio = 0.24;
  const maxFontSize = Math.min(
    72,
    Math.max(6, (height - verticalPadding * 2) / (glyphAscentRatio + glyphDescentRatio)),
  );
  const fontSize = clamp(labelFontSize(spec.fontSize), 6, maxFontSize);
  const ascent = fontSize * glyphAscentRatio;
  const descent = fontSize * glyphDescentRatio;
  const minBaseline = verticalPadding + ascent;
  const maxBaseline = Math.max(minBaseline, height - verticalPadding - descent);
  const y = clamp(finite(anchor.y, minBaseline), minBaseline, maxBaseline);
  const textAnchor = anchor.textAnchor ?? 'middle';
  const keySize = spec.showLegendKey ? clamp(fontSize * 0.45, 4, 12) : 0;
  const keyGap = keySize > 0 ? Math.max(3, fontSize * 0.12) : 0;
  const keyReservation = keySize + keyGap;
  const minimumTextWidth = fullText.length > 0 ? textWidth('…', fontSize) : 0;
  const rawX = finite(anchor.x, width / 2);
  // Like Excel, a label that would run off the chart is moved back inside it
  // whole; only a label wider than the room left is shortened.
  const fitWidth = Math.min(
    textWidth(fullText, fontSize),
    Math.max(minimumTextWidth, width - horizontalPadding * 2 - keyReservation),
  );
  const keptWidth = fullText.length > 0 ? Math.max(minimumTextWidth, fitWidth) : 0;
  const minX =
    textAnchor === 'start'
      ? horizontalPadding + keyReservation
      : textAnchor === 'end'
        ? horizontalPadding + keptWidth
        : horizontalPadding + keyReservation + keptWidth / 2;
  const maxX =
    textAnchor === 'start'
      ? width - horizontalPadding - keptWidth
      : textAnchor === 'end'
        ? width - horizontalPadding - keyReservation
        : width - horizontalPadding - keptWidth / 2;
  const x = clamp(rawX, Math.min(minX, maxX), Math.max(minX, maxX));
  const sideRoom =
    textAnchor === 'start'
      ? width - horizontalPadding - x
      : textAnchor === 'end'
        ? x - horizontalPadding
        : Math.min(x - horizontalPadding - keyReservation, width - horizontalPadding - x) * 2;
  const displayText = truncateLabel(fullText, Math.max(minimumTextWidth, sideRoom), fontSize);
  const attrs: Record<string, string | number> = {
    x,
    y,
    'text-anchor': textAnchor,
    'font-size': fontSize,
    fill: spec.color ?? 'currentColor',
    'aria-label': fullText,
    'white-space': 'pre',
  };
  if (spec.showLegendKey) {
    const displayWidth = textWidth(displayText, fontSize);
    const keyX =
      fullText.length === 0
        ? x - keySize / 2
        : textAnchor === 'end'
          ? x + keyGap
          : textAnchor === 'start'
            ? x - keyGap - keySize
            : x - displayWidth / 2 - keyGap - keySize;
    const keyY = clamp(
      y - keySize * 0.72,
      verticalPadding,
      Math.max(verticalPadding, height - verticalPadding - keySize),
    );
    elements.push(
      primitive(
        'rect',
        {
          x: clamp(keyX, horizontalPadding, Math.max(horizontalPadding, width - keySize)),
          y: keyY,
          width: keySize,
          height: keySize,
          fill: color,
          'aria-label': seriesId,
        },
        'data-label-key',
        seriesIndex,
        pointIndex,
        undefined,
        { seriesId, pointId },
      ),
    );
  }
  if (fullText.length > 0) {
    elements.push(
      primitive('text', attrs, 'data-label', seriesIndex, pointIndex, displayText, {
        seriesId,
        pointId,
      }),
    );
  }
};

interface LabelBounds {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

interface DoughnutLabelRecord {
  readonly seriesIndex: number;
  readonly pointIndex: number;
  readonly originalSeriesIndex: number;
  readonly originalPointIndex: number;
  readonly seriesId: string;
  readonly pointId: string | undefined;
  readonly start: number;
  readonly angle: number;
  readonly midpointRadius: number;
  readonly text: string;
  readonly color: string;
  readonly labelSpec: NonNullable<ChartSpec['dataLabels']>;
}

interface DoughnutLabelCandidate {
  readonly anchor: LabelAnchor;
  readonly bounds: LabelBounds | null;
}

const labelBoundsFor = (elements: readonly ChartPrimitive[]): LabelBounds | null => {
  let bounds: LabelBounds | null = null;
  const include = (next: LabelBounds): void => {
    const padded = {
      left: next.left - 2,
      right: next.right + 2,
      top: next.top - 2,
      bottom: next.bottom + 2,
    };
    bounds =
      bounds === null
        ? padded
        : {
            left: Math.min(bounds.left, padded.left),
            right: Math.max(bounds.right, padded.right),
            top: Math.min(bounds.top, padded.top),
            bottom: Math.max(bounds.bottom, padded.bottom),
          };
  };
  for (const element of elements) {
    if (element.role === 'data-label') {
      const x = Number(element.attrs.x);
      const y = Number(element.attrs.y);
      const fontSize = finite(Number(element.attrs['font-size']), 10);
      if (!Number.isFinite(x) || !Number.isFinite(y) || !(fontSize > 0)) continue;
      const text = element.text ?? '';
      const lines = text.split('\n');
      const textWidthValue = lines.reduce(
        (current, line) => Math.max(current, textWidth(line, fontSize)),
        0,
      );
      const textAnchor = element.attrs['text-anchor'];
      const left =
        textAnchor === 'start'
          ? x
          : textAnchor === 'end'
            ? x - textWidthValue
            : x - textWidthValue / 2;
      include({
        left,
        right: left + textWidthValue,
        top: y - fontSize * 0.99,
        bottom: y + fontSize * 0.24 + Math.max(0, lines.length - 1) * fontSize * 1.1,
      });
    } else if (element.role === 'data-label-key') {
      const x = Number(element.attrs.x);
      const y = Number(element.attrs.y);
      const width = Number(element.attrs.width);
      const height = Number(element.attrs.height);
      if (
        Number.isFinite(x) &&
        Number.isFinite(y) &&
        Number.isFinite(width) &&
        Number.isFinite(height)
      ) {
        include({ left: x, right: x + width, top: y, bottom: y + height });
      }
    }
  }
  return bounds;
};

const labelBoundsOverlap = (left: LabelBounds, right: LabelBounds): boolean =>
  left.left < right.right &&
  right.left < left.right &&
  left.top < right.bottom &&
  right.top < left.bottom;

const addCategoryDataLabels = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  domain: Domain,
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  numericLogBase?: number,
): void => {
  const data = view.data;
  const grouping = spec.grouping ?? 'clustered';
  const segments = segmentsFor(data, grouping);
  const isBar = spec.kind === 'bar';
  const count = Math.max(1, data.categories.length);
  const slot = (isBar ? plot.height : plot.width) / count;
  const numericReverse = isBar ? reversedAxis(spec.xAxis) : reversedAxis(spec.yAxis);
  const categoryReverse = isBar ? reversedAxis(spec.yAxis) : reversedAxis(spec.xAxis);
  const valueScale = (value: number): number =>
    isBar
      ? numericProjection(
          value,
          domain,
          plot.x,
          plot.x + plot.width,
          numericReverse,
          numericLogBase,
        )
      : numericProjection(
          value,
          domain,
          plot.y + plot.height,
          plot.y,
          numericReverse,
          numericLogBase,
        );
  const geometry = barGeometry(slot, Math.max(1, data.series.length), grouping, spec);
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const labelSpec = labelSpecFor(spec, series);
    if (!labelSpec || labelSpec.visible === false) continue;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < data.categories.length; pointIndex += 1) {
      const raw = series.values[pointIndex] ?? null;
      if (raw === null || !Number.isFinite(raw) || (numericLogBase !== undefined && !(raw > 0)))
        continue;
      const segment = segments[seriesIndex]?.[pointIndex];
      if (!segment) continue;
      const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
      const pointId = categoryIdAt(view, pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      const categoryName = data.categories[pointIndex];
      const percentage = grouping === 'percent-stacked' ? Math.abs(segment.value) : undefined;
      const fullText = labelParts(
        labelSpec,
        series.name ?? String(originalSeriesIndex + 1),
        categoryName,
        raw,
        percentage,
        locale,
        formatter,
      );
      if (fullText.length === 0 && !labelSpec.showLegendKey) continue;
      const valueStart = valueScale(
        numericLogBase !== undefined && !(segment.start > 0) ? domain.min : segment.start,
      );
      const valueEnd = valueScale(segment.end);
      const valueDelta = valueEnd - valueStart;
      const rawPositive = raw >= 0;
      const screenPositive = isBar
        ? valueDelta !== 0
          ? valueDelta > 0
          : rawPositive !== numericReverse
        : valueDelta !== 0
          ? valueDelta < 0
          : rawPositive !== numericReverse;
      let anchor: LabelAnchor;
      const requested = resolveChartLabelPosition(spec, labelSpec.position);
      if (spec.kind === 'column' || spec.kind === 'bar') {
        const categoryPosition = categoryIndex(count, pointIndex, categoryReverse);
        const x =
          spec.kind === 'bar'
            ? Math.min(valueStart, valueEnd)
            : plot.x +
              categoryPosition * slot +
              geometry.groupOffset +
              seriesIndex * geometry.seriesStep +
              geometry.size / 2;
        const y =
          spec.kind === 'bar'
            ? plot.y +
              categoryPosition * slot +
              geometry.groupOffset +
              seriesIndex * geometry.seriesStep +
              geometry.size / 2
            : Math.min(valueStart, valueEnd);
        if (spec.kind === 'bar') {
          const endX = screenPositive
            ? Math.max(valueStart, valueEnd)
            : Math.min(valueStart, valueEnd);
          // For stacked bars, the segment's start is its visible base. Using
          // zero here puts every inside-base label on the axis instead.
          const baseX = valueStart;
          if (requested === 'center') {
            anchor = {
              x: (Math.min(valueStart, valueEnd) + Math.max(valueStart, valueEnd)) / 2,
              y,
              textAnchor: 'middle',
            };
          } else if (requested === 'inside-base') {
            anchor = {
              x: screenPositive ? baseX + 4 : baseX - 4,
              y,
              textAnchor: screenPositive ? 'start' : 'end',
            };
          } else if (requested === 'inside-end' || requested === 'best-fit') {
            anchor = {
              x: screenPositive ? endX - 4 : endX + 4,
              y,
              textAnchor: screenPositive ? 'end' : 'start',
            };
          } else {
            anchor = {
              x: screenPositive ? endX + 4 : endX - 4,
              y,
              textAnchor: screenPositive ? 'start' : 'end',
            };
          }
        } else if (requested === 'center') {
          anchor = { x, y: (Math.min(valueStart, valueEnd) + Math.max(valueStart, valueEnd)) / 2 };
        } else if (requested === 'inside-base') {
          // For stacked columns, the segment's start is its visible base.
          const baseY = valueStart;
          anchor = { x, y: screenPositive ? baseY - 4 : baseY + 12 };
        } else if (requested === 'inside-end' || requested === 'best-fit') {
          const endY = screenPositive
            ? Math.min(valueStart, valueEnd)
            : Math.max(valueStart, valueEnd);
          anchor = { x, y: screenPositive ? endY + 12 : endY - 4 };
        } else if (requested === 'left') {
          anchor = { x: x - 5, y, textAnchor: 'end' };
        } else if (requested === 'right') {
          anchor = { x: x + 5, y, textAnchor: 'start' };
        } else if (requested === 'below') {
          anchor = { x, y: Math.max(valueStart, valueEnd) + 12 };
        } else {
          const endY = screenPositive
            ? Math.min(valueStart, valueEnd)
            : Math.max(valueStart, valueEnd);
          anchor = { x, y: screenPositive ? endY - 4 : endY + 12 };
        }
      } else {
        anchor = {
          x: xCategory(plot, data.categories.length, pointIndex, categoryReverse),
          y: valueEnd - 5,
        };
        if (requested === 'below') anchor = { ...anchor, y: valueEnd + 12 };
        if (requested === 'left') anchor = { ...anchor, x: anchor.x - 5, textAnchor: 'end' };
        if (requested === 'right') anchor = { ...anchor, x: anchor.x + 5, textAnchor: 'start' };
        if (requested === 'center') anchor = { ...anchor, y: valueEnd };
      }
      addDataLabel(
        elements,
        fullText,
        anchor,
        labelSpec,
        width,
        height,
        originalSeriesIndex,
        originalPointIndex,
        series.id,
        pointId,
        keyColor(spec, style),
      );
    }
  }
};

const addScatterDataLabels = (
  elements: ChartPrimitive[],
  view: XYChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  xDomain: Domain,
  yDomain: Domain,
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  xLogBase?: number,
  yLogBase?: number,
): void => {
  const data = view.data;
  const xScale = (value: number): number =>
    numericProjection(
      value,
      xDomain,
      plot.x,
      plot.x + plot.width,
      reversedAxis(spec.xAxis),
      xLogBase,
    );
  const yScale = (value: number): number =>
    numericProjection(
      value,
      yDomain,
      plot.y + plot.height,
      plot.y,
      reversedAxis(spec.yAxis),
      yLogBase,
    );
  const anyLogAxis = xLogBase !== undefined || yLogBase !== undefined;
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const labelSpec = labelSpecFor(spec, series);
    if (!labelSpec || labelSpec.visible === false) continue;
    // The X part follows the x values' format unless the labels set their own.
    const ownCode = labelSpecFor(spec, { id: series.id })?.numberFormat;
    const xLabelCode =
      ownCode !== undefined && ownCode.trim().length > 0 ? ownCode : series.xFormatCode;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < series.points.length; pointIndex += 1) {
      const point = series.points[pointIndex];
      if (!point) continue;
      if (
        point.x === null ||
        point.y === null ||
        !Number.isFinite(point.x) ||
        !Number.isFinite(point.y) ||
        (xLogBase !== undefined && !(point.x > 0)) ||
        (yLogBase !== undefined && !(point.y > 0)) ||
        (anyLogAxis &&
          (point.x < xDomain.min ||
            point.x > xDomain.max ||
            point.y < yDomain.min ||
            point.y > yDomain.max))
      )
        continue;
      const originalPointIndex = view.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex;
      const pointId = pointIdAt(view, seriesIndex, pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      const fullText = labelParts(
        labelSpec,
        series.name ?? String(originalSeriesIndex + 1),
        formatChartValue(point.x, xLabelCode, locale, formatter) ?? undefined,
        point.y,
        undefined,
        locale,
        formatter,
      );
      const x = xScale(point.x);
      const y = yScale(point.y);
      const requested = resolveChartLabelPosition(spec, labelSpec.position);
      const anchor: LabelAnchor =
        requested === 'center'
          ? { x, y }
          : requested === 'below'
            ? { x, y: y + 12 }
            : requested === 'left'
              ? { x: x - 5, y, textAnchor: 'end' }
              : requested === 'right'
                ? { x: x + 5, y, textAnchor: 'start' }
                : { x, y: y - 5 };
      addDataLabel(
        elements,
        fullText,
        anchor,
        labelSpec,
        width,
        height,
        originalSeriesIndex,
        originalPointIndex,
        series.id,
        pointId,
        keyColor(spec, style),
      );
    }
  }
};

const addPieDataLabels = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): void => {
  const data = view.data;
  const first = data.series[0];
  if (!first) return;
  const labelSpec = labelSpecFor(spec, first);
  if (!labelSpec || labelSpec.visible === false) return;
  const magnitudes = first.values.map((value) =>
    value !== null && Number.isFinite(value) ? Math.abs(value) : 0,
  );
  const scale = magnitudes.reduce((maximum, value) => Math.max(maximum, value), 0);
  if (!(scale > 0)) return;
  const normalized = magnitudes.map((value) => value / scale);
  const total = normalized.reduce((sum, value) => sum + value, 0);
  if (!(total > 0)) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const radius = Math.max(0, Math.min(plot.width, plot.height) / 2 - 2);
  let start = -Math.PI / 2 + (boundedOption(spec.firstSliceAngle, 0, 360, 0) * Math.PI) / 180;
  for (let pointIndex = 0; pointIndex < normalized.length; pointIndex += 1) {
    const normalizedValue = normalized[pointIndex] ?? 0;
    if (!(normalizedValue > 0)) continue;
    const angle = (normalizedValue / total) * Math.PI * 2;
    const middle = start + angle / 2;
    const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
    const pointId = categoryIdAt(view, pointIndex);
    const style = resolveChartStyle(
      spec,
      palette,
      first.id,
      pointId,
      view.originalSeriesIndices[0] ?? 0,
      originalPointIndex,
    );
    const raw = first.values[pointIndex];
    if (raw === null || !Number.isFinite(raw)) {
      start += angle;
      continue;
    }
    const fullText = labelParts(
      labelSpec,
      first.name ?? String((view.originalSeriesIndices[0] ?? 0) + 1),
      data.categories[pointIndex],
      raw,
      normalizedValue / total,
      locale,
      formatter,
    );
    const requested = resolveChartLabelPosition(spec, labelSpec.position);
    const distance =
      requested === 'outside-end'
        ? radius + 8
        : requested === 'inside-end'
          ? radius * 0.85
          : radius * 0.58;
    const rawAnchor = {
      x: cx + Math.cos(middle) * distance,
      y: cy + Math.sin(middle) * distance,
    };
    const anchor: LabelAnchor =
      requested === 'above'
        ? { ...rawAnchor, y: rawAnchor.y - radius * 0.2 }
        : requested === 'below'
          ? { ...rawAnchor, y: rawAnchor.y + radius * 0.2 }
          : requested === 'left'
            ? { ...rawAnchor, x: rawAnchor.x - radius * 0.2, textAnchor: 'end' }
            : requested === 'right'
              ? { ...rawAnchor, x: rawAnchor.x + radius * 0.2, textAnchor: 'start' }
              : {
                  ...rawAnchor,
                  textAnchor:
                    requested === 'outside-end'
                      ? Math.cos(middle) > 0.2
                        ? 'start'
                        : Math.cos(middle) < -0.2
                          ? 'end'
                          : 'middle'
                      : 'middle',
                };
    addDataLabel(
      elements,
      fullText,
      anchor,
      labelSpec,
      width,
      height,
      view.originalSeriesIndices[0] ?? 0,
      originalPointIndex,
      first.id,
      pointId,
      fillColor(style.fill, style.color),
    );
    start += angle;
  }
};

const addDoughnutDataLabels = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): void => {
  const data = view.data;
  if (data.series.length === 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const outerRadius = Math.max(0, Math.min(plot.width, plot.height) / 2 - 2);
  const holePercent = boundedOption(spec.holeSize, 10, 90, 50);
  const holeRadius = (outerRadius * holePercent) / 100;
  const ringWidth = (outerRadius - holeRadius) / data.series.length;
  const firstSliceAngle = boundedOption(spec.firstSliceAngle, 0, 360, 0);
  const labelAngles = (start: number, angle: number): readonly number[] => {
    const middle = start + angle / 2;
    if (angle >= Math.PI * 2 - 1e-9) {
      return [
        middle,
        middle + Math.PI,
        middle + Math.PI / 2,
        middle - Math.PI / 2,
        middle + Math.PI / 4,
        middle - Math.PI / 4,
        middle + (Math.PI * 3) / 4,
        middle - (Math.PI * 3) / 4,
      ];
    }
    const fractions = [0, -0.15, 0.15, -0.3, 0.3, -0.45, 0.45, -0.49, 0.49];
    const minimum = start + angle * 0.02;
    const maximum = start + angle * 0.98;
    return fractions.map((fraction) => clamp(middle + fraction * angle, minimum, maximum));
  };
  const records: DoughnutLabelRecord[] = [];
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const labelSpec = labelSpecFor(spec, series);
    if (!labelSpec || labelSpec.visible === false) continue;
    const normalized = normalizedMagnitudes(series.values);
    if (!normalized) continue;
    const innerRadius = holeRadius + ringWidth * seriesIndex;
    const outerRingRadius = innerRadius + ringWidth;
    const midpointRadius = (innerRadius + outerRingRadius) / 2;
    let start = -Math.PI / 2 + (firstSliceAngle * Math.PI) / 180;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < normalized.values.length; pointIndex += 1) {
      const normalizedValue = normalized.values[pointIndex] ?? 0;
      if (!(normalizedValue > 0) || !Number.isFinite(normalizedValue)) continue;
      const angle = (normalizedValue / normalized.total) * Math.PI * 2;
      if (!(angle > 0) || !Number.isFinite(angle)) continue;
      const raw = series.values[pointIndex];
      if (raw === null || !Number.isFinite(raw)) {
        start += angle;
        continue;
      }
      const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
      const pointId = categoryIdAt(view, pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      const fullText = labelParts(
        labelSpec,
        series.name ?? String(originalSeriesIndex + 1),
        data.categories[pointIndex],
        raw,
        normalizedValue / normalized.total,
        locale,
        formatter,
      );
      records.push({
        seriesIndex,
        pointIndex,
        originalSeriesIndex,
        originalPointIndex,
        seriesId: series.id,
        pointId,
        start,
        angle,
        midpointRadius,
        text: fullText,
        color: fillColor(style.fill, style.color),
        labelSpec,
      });
      start += angle;
    }
  }

  const candidatesFor = new Map<DoughnutLabelRecord, readonly DoughnutLabelCandidate[]>();
  for (const record of records) {
    const candidates = labelAngles(record.start, record.angle).map((candidateAngle) => {
      const candidateAnchor: LabelAnchor = {
        ...circularPoint(cx, cy, record.midpointRadius, candidateAngle),
        textAnchor: 'middle',
      };
      const candidateElements: ChartPrimitive[] = [];
      addDataLabel(
        candidateElements,
        record.text,
        candidateAnchor,
        record.labelSpec,
        width,
        height,
        record.originalSeriesIndex,
        record.originalPointIndex,
        record.seriesId,
        record.pointId,
        record.color,
      );
      return { anchor: candidateAnchor, bounds: labelBoundsFor(candidateElements) };
    });
    candidatesFor.set(record, candidates);
  }

  const placementOrder = records
    .slice()
    .sort((left, right) =>
      left.pointIndex === right.pointIndex
        ? left.seriesIndex - right.seriesIndex
        : left.pointIndex - right.pointIndex,
    );
  const selections = new Map<DoughnutLabelRecord, DoughnutLabelCandidate>();
  const placedLabels: LabelBounds[] = [];
  const fitsPlacedLabels = (candidate: DoughnutLabelCandidate): boolean => {
    const bounds = candidate.bounds;
    if (bounds === null) return true;
    return placedLabels.every((placed) => !labelBoundsOverlap(bounds, placed));
  };
  const maxPlacementStates = 4096;
  let placementStates = 0;
  const tryPlace = (index: number): boolean => {
    if (index >= placementOrder.length) return true;
    placementStates += 1;
    if (placementStates > maxPlacementStates) return false;
    const record = placementOrder[index];
    if (!record) return false;
    const candidates = candidatesFor.get(record) ?? [];
    for (const candidate of candidates) {
      if (!fitsPlacedLabels(candidate)) continue;
      selections.set(record, candidate);
      if (candidate.bounds !== null) placedLabels.push(candidate.bounds);
      if (tryPlace(index + 1)) return true;
      if (candidate.bounds !== null) placedLabels.pop();
      selections.delete(record);
    }
    return false;
  };
  const placedWithoutBacktracking = placementOrder.length <= 64 && tryPlace(0);
  if (!placedWithoutBacktracking) {
    selections.clear();
    placedLabels.length = 0;
    for (const record of placementOrder) {
      const candidates = candidatesFor.get(record) ?? [];
      const selected = candidates.find((candidate) => fitsPlacedLabels(candidate)) ?? candidates[0];
      if (!selected) continue;
      selections.set(record, selected);
      if (selected.bounds !== null) placedLabels.push(selected.bounds);
    }
  }

  for (const record of records) {
    const selection = selections.get(record);
    if (!selection) continue;
    addDataLabel(
      elements,
      record.text,
      selection.anchor,
      record.labelSpec,
      width,
      height,
      record.originalSeriesIndex,
      record.originalPointIndex,
      record.seriesId,
      record.pointId,
      record.color,
    );
  }
};

const addDataLabels = (
  elements: ChartPrimitive[],
  view: PreparedChartData,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  xDomain: Domain | null,
  yDomain: Domain | null,
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  categoryLogBase?: number,
  scatterLogBases?: { readonly x?: number; readonly y?: number },
): void => {
  if (!hasDataLabels(spec)) return;
  if (view.data.domain === 'category') {
    if (spec.kind === 'pie')
      addPieDataLabels(
        elements,
        view as CategoryChartView,
        spec,
        plot,
        palette,
        width,
        height,
        locale,
        formatter,
      );
    else if (spec.kind === 'doughnut')
      addDoughnutDataLabels(
        elements,
        view as CategoryChartView,
        spec,
        plot,
        palette,
        width,
        height,
        locale,
        formatter,
      );
    else if (yDomain || spec.kind === 'bar')
      addCategoryDataLabels(
        elements,
        view as CategoryChartView,
        spec,
        plot,
        (spec.kind === 'bar' ? xDomain : yDomain) as Domain,
        palette,
        width,
        height,
        locale,
        formatter,
        categoryLogBase,
      );
  } else if (xDomain && yDomain) {
    addScatterDataLabels(
      elements,
      view as XYChartView,
      spec,
      plot,
      xDomain,
      yDomain,
      palette,
      width,
      height,
      locale,
      formatter,
      scatterLogBases?.x,
      scatterLogBases?.y,
    );
  }
};

/**
 * Convert chart data into a finite, browser-independent scene. The default
 * 320×132 scene retains the historic 20/20/300/112 plot rectangle; title and
 * legend reservations are added only when those features are requested.
 */
export function layoutChart(input: {
  readonly spec: ChartSpec;
  readonly data: ChartData;
  readonly width: number;
  readonly height: number;
  readonly locale?: string;
  readonly formatValue?: ChartValueFormatter;
  /** The workbook theme's accent1..accent6; the default Office theme's when omitted. */
  readonly accents?: readonly string[];
  /** The automatic title of a chart with several series, localised; "Chart Title" when omitted. */
  readonly chartTitle?: string;
  /** The automatic title of an axis, localised; "Axis Title" when omitted. */
  readonly axisTitle?: string;
}): ChartScene {
  const width = Math.max(1, finite(input.width, 320));
  const height = Math.max(1, finite(input.height, 132));
  const maxNumericLabelWidth = numericAxisLabelMaxWidth(width);
  const locale = input.locale ?? 'en-US';
  const prepared = withCategoryFormat(
    withPointStates(prepareChartData(input.data, input.spec), input.data),
    input.data,
    locale,
    input.formatValue,
  );
  const spec = withValueAxisFormat(input.spec, prepared);
  const data = prepared.data;
  const categoryXLogBase =
    data.domain === 'category' ? categoryNumericLogBase(spec, 'x') : undefined;
  const categoryYLogBase =
    data.domain === 'category' ? categoryNumericLogBase(spec, 'y') : undefined;
  const scatterLogs = data.domain === 'xy' && spec.kind === 'scatter' ? scatterLogBases(spec) : {};
  const xLogBase = data.domain === 'category' ? categoryXLogBase : scatterLogs.x;
  const yLogBase = data.domain === 'category' ? categoryYLogBase : scatterLogs.y;
  const palette = paletteFor(spec, input.accents);
  const scatterPointEligible = (
    point: XYChartData['series'][number]['points'][number],
  ): boolean => {
    return (
      point.x !== null &&
      point.y !== null &&
      Number.isFinite(point.x) &&
      Number.isFinite(point.y) &&
      (scatterLogs.x === undefined || point.x > 0) &&
      (scatterLogs.y === undefined || point.y > 0)
    );
  };
  const elements: ChartPrimitive[] = [
    primitive('rect', { x: 0, y: 0, width, height, fill: 'transparent' }, 'background'),
  ];
  const hasLegend =
    spec.legend?.visible ??
    (data.domain === 'category'
      ? isCircularKind(spec.kind)
        ? data.categories.length > 1 && (spec.kind === 'pie' || data.series.length > 0)
        : data.series.length > 1
      : data.series.length > 1);
  const legendPosition = spec.legend?.position ?? 'bottom';
  const title = chartTitleText(spec.title, input.data, input.chartTitle);
  let left = 20;
  let right = 20;
  let top = title ? TITLE_RESERVE : 20;
  let bottom = 20;
  let rightLegendReserve = 0;
  if (hasLegend) {
    const legendLabels =
      data.domain === 'category' && isCircularKind(spec.kind)
        ? data.categories
        : data.series.map((series, index) => series.name ?? `${index + 1}`);
    const legendWidth = legendLabels.reduce(
      (current, label) => Math.max(current, textWidth(label) + 20),
      56,
    );
    if (legendPosition === 'right') {
      rightLegendReserve = Math.min(180, Math.max(56, legendWidth));
      right += rightLegendReserve;
    }
    if (legendPosition === 'left') left += Math.min(100, Math.max(56, width * 0.24));
    if (legendPosition === 'top') top = Math.max(top, spec.title ? TITLE_RESERVE + 14 : 20);
    if (legendPosition === 'bottom') {
      // Category labels and an optional x-axis title occupy the space above
      // a bottom legend. Reserve enough rows for both at compact sizes.
      bottom = Math.max(
        bottom,
        !isCircularKind(spec.kind) && axisTitleText(spec.xAxis?.title) && axisIsVisible(spec.xAxis)
          ? 40
          : 30,
      );
    }
  }
  const highXLabels =
    !isCircularKind(spec.kind) &&
    axisIsVisible(spec.xAxis) &&
    resolvedAxisLabelPosition(spec.xAxis) === 'high';
  if (highXLabels && (Boolean(spec.title) || (hasLegend && legendPosition === 'top'))) {
    top += 14;
  }
  const highYLabels =
    !isCircularKind(spec.kind) &&
    axisIsVisible(spec.yAxis) &&
    resolvedAxisLabelPosition(spec.yAxis) === 'high';
  let rightLabelGutter = 0;
  if (spec.kind === 'bar' && data.domain === 'category' && axisIsVisible(spec.yAxis)) {
    const categoryWidth = data.categories.reduce(
      (current, label) =>
        Math.max(
          current,
          textWidth(highYLabels ? boundedAxisLabel(label, maxNumericLabelWidth) : label) + 10,
        ),
      20,
    );
    if (highYLabels) {
      rightLabelGutter = Math.max(rightLabelGutter, categoryWidth - 2);
      if (!(hasLegend && legendPosition === 'right')) {
        right = Math.max(right, rightLabelGutter);
      }
    } else {
      left = Math.max(left, Math.min(Math.max(20, width - 8), categoryWidth));
    }
  }
  const yValueCandidates =
    data.domain === 'category'
      ? (() => {
          const grouping = spec.grouping ?? 'clustered';
          if (grouping === 'clustered') {
            return data.series.flatMap((series) =>
              series.values.flatMap((value) => {
                const resolved = valueFor(value);
                return resolved === null ? [] : [resolved];
              }),
            );
          }
          const totals = effectiveStackTotals(data);
          return [...totals.positive, ...totals.negative];
        })()
      : data.series.flatMap((series) =>
          series.points.flatMap((point) => {
            if (!scatterPointEligible(point)) return [];
            return point.y === null ? [] : [point.y];
          }),
        );
  const xValueCandidates =
    spec.kind === 'bar' && data.domain === 'category'
      ? yValueCandidates
      : spec.kind === 'scatter' && data.domain === 'xy'
        ? data.series.flatMap((series) =>
            series.points.flatMap((point) => {
              if (!scatterPointEligible(point)) return [];
              return point.x === null ? [] : [point.x];
            }),
          )
        : [];
  const xEstimate =
    xValueCandidates.length > 0 || xLogBase !== undefined
      ? axisDomain(
          xValueCandidates,
          spec.xAxis?.min,
          spec.xAxis?.max,
          spec.kind === 'bar' && spec.grouping === 'percent-stacked',
          xLogBase,
        )
      : null;
  if (
    xEstimate &&
    axisIsVisible(spec.xAxis) &&
    (xLogBase !== undefined ||
      (spec.xAxis?.numberFormat !== undefined && spec.xAxis.numberFormat.trim().length > 0))
  ) {
    const xLabels = axisTicks(xEstimate, spec.xAxis, xLogBase).map((tick) =>
      boundedAxisLabel(
        axisTickLabel(
          tick,
          spec.kind === 'bar' && spec.grouping === 'percent-stacked',
          spec.xAxis,
          input.locale ?? 'en-US',
          input.formatValue,
        ),
        maxNumericLabelWidth,
      ),
    );
    const firstWidth = xLabels.length > 0 ? textWidth(xLabels[0] ?? '') : 0;
    const lastWidth = xLabels.length > 0 ? textWidth(xLabels[xLabels.length - 1] ?? '') : 0;
    const leftEndpointWidth = reversedAxis(spec.xAxis) ? lastWidth : firstWidth;
    const rightEndpointWidth = reversedAxis(spec.xAxis) ? firstWidth : lastWidth;
    // Numeric x labels are centered on their tick. Reserve each endpoint's
    // half-width so a long host format cannot clip against the canvas edge.
    const maxMargin = Math.max(20, width - 8);
    left = Math.max(left, Math.min(maxMargin, leftEndpointWidth / 2 + 2));
    right = Math.max(right, Math.min(maxMargin, rightEndpointWidth / 2 + 2));
  }
  let yLabelWidth =
    spec.kind === 'bar' && data.domain === 'category' && axisIsVisible(spec.yAxis)
      ? data.categories.reduce(
          (current, label) =>
            Math.max(
              current,
              textWidth(highYLabels ? boundedAxisLabel(label, maxNumericLabelWidth) : label),
            ),
          12,
        )
      : 12;
  if (
    spec.kind !== 'bar' &&
    !isCircularKind(spec.kind) &&
    (yValueCandidates.length > 0 || yLogBase !== undefined) &&
    axisIsVisible(spec.yAxis)
  ) {
    const yEstimate = axisDomain(
      yValueCandidates,
      spec.yAxis?.min,
      spec.yAxis?.max,
      data.domain === 'category' && spec.grouping === 'percent-stacked',
      yLogBase,
    );
    yLabelWidth = axisTicks(yEstimate, spec.yAxis, yLogBase).reduce(
      (current, tick) =>
        Math.max(
          current,
          textWidth(
            boundedAxisLabel(
              axisTickLabel(
                tick,
                data.domain === 'category' && spec.grouping === 'percent-stacked',
                spec.yAxis,
                input.locale ?? 'en-US',
                input.formatValue,
              ),
              maxNumericLabelWidth,
            ),
          ),
        ),
      12,
    );
    if (yLabelWidth > 18) {
      // SVG text metrics are a little wider than the character estimate on
      // the default system font. Keep a small safety gutter so the first
      // signed/percent tick remains inside the canvas.
      if (highYLabels) {
        rightLabelGutter = Math.max(rightLabelGutter, yLabelWidth + 8);
        if (!(hasLegend && legendPosition === 'right')) {
          right = Math.max(right, rightLabelGutter);
        }
      } else {
        left = Math.max(left, Math.min(Math.max(20, width - 8), yLabelWidth + 6));
      }
    }
  }
  if (!isCircularKind(spec.kind) && axisIsVisible(spec.yAxis) && axisTitleText(spec.yAxis?.title)) {
    // A rotated title still needs a horizontal gutter beside wide y tick
    // labels. Keep the historic 36px minimum for ordinary labels, then add
    // a separate reserve once labels exceed that width.
    left = Math.max(left, highYLabels ? 36 : yLabelWidth > 36 ? yLabelWidth + 32 : 36);
  }
  if (!isCircularKind(spec.kind) && axisIsVisible(spec.xAxis) && axisTitleText(spec.xAxis?.title))
    bottom = Math.max(bottom, 32);
  if (highYLabels) {
    rightLabelGutter = Math.max(rightLabelGutter, yLabelWidth + 8);
    if (hasLegend && legendPosition === 'right') right += rightLabelGutter;
    else right = Math.max(right, rightLabelGutter);
  }
  left = Math.min(left, Math.max(1, width - 2));
  right = Math.min(right, Math.max(1, width - left - 1));
  top = Math.min(top, Math.max(1, height - 2));
  bottom = Math.min(bottom, Math.max(1, height - top - 1));
  const plot = {
    x: left,
    y: top,
    width: Math.max(1, width - left - right),
    height: Math.max(1, height - top - bottom),
  };
  const allocatedRightLabelGutter = highYLabels
    ? Math.max(0, Math.min(rightLabelGutter, right - rightLegendReserve))
    : 0;
  const axisLabelWidth =
    highYLabels && allocatedRightLabelGutter > 0
      ? Math.min(maxNumericLabelWidth, Math.max(12, allocatedRightLabelGutter - 6))
      : maxNumericLabelWidth;
  if (title) {
    elements.push(
      primitive(
        'text',
        { x: width / 2, y: TITLE_BASELINE, 'text-anchor': 'middle' },
        'title',
        undefined,
        undefined,
        title,
      ),
    );
  }

  let xDomain: Domain | null = null;
  let yDomain: Domain | null = null;
  let labels: readonly string[] = [];
  let percent = false;
  if (data.domain === 'category') {
    labels = data.categories;
    const grouping = spec.grouping ?? 'clustered';
    percent = grouping === 'percent-stacked';
    const totals = effectiveStackTotals(data);
    const rawValues = data.series.flatMap((series) =>
      series.values.flatMap((value) => {
        const resolved = valueFor(value);
        return resolved === null ? [] : [resolved];
      }),
    );
    const stackValues =
      grouping === 'clustered' ? rawValues : [...totals.positive, ...totals.negative];
    const horizontal = spec.kind === 'bar';
    const axisSpec = horizontal ? spec.xAxis : spec.yAxis;
    const domain = axisDomain(
      stackValues,
      axisSpec?.min,
      axisSpec?.max,
      percent,
      horizontal ? categoryXLogBase : categoryYLogBase,
    );
    if (horizontal) xDomain = domain;
    else yDomain = domain;
    if (!isCircularKind(spec.kind)) {
      addCategoryMarks(
        elements,
        prepared as CategoryChartView,
        spec,
        plot,
        domain,
        palette,
        horizontal ? categoryXLogBase : categoryYLogBase,
      );
      addAxis(
        elements,
        plot,
        spec.kind,
        xDomain,
        yDomain,
        labels,
        (prepared as CategoryChartView).originalCategoryIndices,
        data.categoryIds ?? [],
        spec,
        percent,
        height,
        maxNumericLabelWidth,
        input.locale ?? 'en-US',
        input.formatValue,
        axisLabelWidth,
        categoryXLogBase,
        categoryYLogBase,
        input.axisTitle,
      );
    }
  } else if (spec.kind === 'scatter') {
    const pairedPoints = data.series.flatMap((series) =>
      series.points.filter(
        (point) =>
          point.x !== null &&
          point.y !== null &&
          Number.isFinite(point.x) &&
          Number.isFinite(point.y) &&
          (scatterLogs.x === undefined || point.x > 0) &&
          (scatterLogs.y === undefined || point.y > 0),
      ),
    );
    const xValues = pairedPoints.flatMap((point) => (point.x === null ? [] : [point.x]));
    const yValues = pairedPoints.flatMap((point) => (point.y === null ? [] : [point.y]));
    xDomain = axisDomain(xValues, spec.xAxis?.min, spec.xAxis?.max, false, scatterLogs.x);
    yDomain = axisDomain(yValues, spec.yAxis?.min, spec.yAxis?.max, false, scatterLogs.y);
    addScatterMarks(
      elements,
      prepared as XYChartView,
      spec,
      plot,
      xDomain,
      yDomain,
      palette,
      scatterLogs.x,
      scatterLogs.y,
    );
    addAxis(
      elements,
      plot,
      spec.kind,
      xDomain,
      yDomain,
      [],
      [],
      [],
      spec,
      false,
      height,
      maxNumericLabelWidth,
      input.locale ?? 'en-US',
      input.formatValue,
      axisLabelWidth,
      scatterLogs.x,
      scatterLogs.y,
      input.axisTitle,
    );
  } else {
    // A malformed domain/kind pair is still laid out safely as an empty chart.
    addAxis(
      elements,
      plot,
      spec.kind,
      null,
      null,
      [],
      [],
      [],
      spec,
      false,
      height,
      maxNumericLabelWidth,
      input.locale ?? 'en-US',
      input.formatValue,
      axisLabelWidth,
      undefined,
      undefined,
      input.axisTitle,
    );
  }

  if (spec.kind === 'pie' && data.domain === 'category') {
    addPieMarks(elements, prepared as CategoryChartView, plot, palette, spec);
  } else if (spec.kind === 'doughnut' && data.domain === 'category') {
    addDoughnutMarks(elements, prepared as CategoryChartView, plot, palette, spec);
  }
  addDataLabels(
    elements,
    prepared,
    spec,
    plot,
    xDomain,
    yDomain,
    palette,
    width,
    height,
    input.locale ?? 'en-US',
    input.formatValue,
    data.domain === 'category'
      ? spec.kind === 'bar'
        ? categoryXLogBase
        : categoryYLogBase
      : undefined,
    data.domain === 'xy' ? scatterLogs : undefined,
  );
  if (hasLegend)
    addLegend(elements, prepared, spec, width, height, plot, palette, allocatedRightLabelGutter);
  if (
    !dataHasValues(
      data,
      spec.kind,
      data.domain === 'category'
        ? spec.kind === 'bar'
          ? categoryXLogBase
          : categoryYLogBase
        : undefined,
      data.domain === 'xy' ? scatterLogs : undefined,
    )
  ) {
    elements.push(
      primitive(
        'text',
        { x: plot.x + plot.width / 2, y: plot.y + plot.height / 2, 'text-anchor': 'middle' },
        'empty',
        undefined,
        undefined,
        '—',
      ),
    );
  }
  return { width, height, plot, elements };
}

/**
 * Scene for a chart the engine models but cannot draw: a frame, the chart
 * title when it has one, and a caller-supplied (localized) message.
 */
export function layoutUnsupportedChart(input: {
  readonly title?: string;
  readonly message: string;
  readonly width: number;
  readonly height: number;
}): ChartScene {
  const width = Math.max(1, finite(input.width, 320));
  const height = Math.max(1, finite(input.height, 132));
  const plot = { x: 0, y: 0, width, height };
  const elements: ChartPrimitive[] = [
    primitive('rect', { x: 0, y: 0, width, height, fill: 'transparent' }, 'background'),
    primitive(
      'rect',
      {
        x: 0.5,
        y: 0.5,
        width: Math.max(0, width - 1),
        height: Math.max(0, height - 1),
        fill: 'none',
        stroke: 'currentColor',
        'stroke-opacity': 0.4,
      },
      'frame',
    ),
  ];
  if (input.title) {
    elements.push(
      primitive(
        'text',
        { x: width / 2, y: TITLE_BASELINE, 'text-anchor': 'middle' },
        'title',
        undefined,
        undefined,
        input.title,
      ),
    );
  }
  elements.push(
    primitive(
      'text',
      { x: width / 2, y: height / 2, 'text-anchor': 'middle' },
      'unsupported',
      undefined,
      undefined,
      input.message,
    ),
  );
  return { width, height, plot, elements };
}
