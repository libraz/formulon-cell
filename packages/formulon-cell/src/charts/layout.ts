import {
  addAxis,
  axisIsVisible,
  axisTickLabel,
  axisTicks,
  axisTitleText,
  boundedAxisLabel,
  numericAxisLabelMaxWidth,
  resolvedAxisLabelPosition,
} from './axes.js';
import { addDataLabels } from './data-labels.js';
import {
  type CategoryChartView,
  type PreparedChartData,
  prepareChartData,
  type XYChartView,
} from './filters.js';
import { formatChartValue } from './format.js';
import { addLegend } from './legend.js';
import {
  axisDomain,
  categoryNumericLogBase,
  type Domain,
  reversedAxis,
  scatterLogBases,
} from './plot-scale.js';
import { primitive, TITLE_BASELINE, TITLE_RESERVE, textWidth } from './scene-primitives.js';
import { automaticPalette } from './series-colors.js';
import { effectiveStackTotals, isCircularKind, valueFor } from './series-geometry.js';
import {
  addCategoryMarks,
  addDoughnutMarks,
  addPieMarks,
  addScatterMarks,
} from './series-marks.js';
import type {
  ChartAxisSpec,
  ChartData,
  ChartKind,
  ChartPrimitive,
  ChartScene,
  ChartSpec,
  ChartValueFormatter,
  XYChartData,
} from './types.js';

export { axisTitleText } from './axes.js';

const finite = (value: number, fallback: number): number =>
  Number.isFinite(value) ? value : fallback;

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

/** The spec's palette, else Excel's automatic colours in the workbook theme. */
const paletteFor = (spec: ChartSpec, accents: readonly string[] | undefined): readonly string[] =>
  spec.palette && spec.palette.length > 0 ? spec.palette : automaticPalette(accents);

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
