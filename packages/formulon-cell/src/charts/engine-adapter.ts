/**
 * Mapping between the engine's chart model and `ChartSpec`.
 *
 * The engine owns the chart; a `ChartSpec` is its projection. `CHART_SPEC_FIELDS`
 * holds one entry per `ChartSpec` field with the kinds it applies to, how it reads
 * from `ChartInfo` and which engine operations a change to it produces. A field
 * outside its kinds reads as undefined and produces no operation.
 *
 * Identities: a series is `series-${c:idx}`, a formatted point `point-${idx}`, a
 * hidden category `category-${i}` (a scatter chart filters by `point-${i}`, the id
 * its points carry). Colours are `#rrggbb` here and ARGB numbers in the engine;
 * transparency is 0–100 on both sides.
 *
 * Operations address axes by role (`x` is the category-role axis, or a scatter
 * chart's horizontal value axis) because `setChartType` replaces the axes; resolve
 * the id with `chartAxisId` against the chart as it is when the operation runs.
 * Everything here is pure: callers execute the operations.
 */

import type {
  ChartAxisInfo,
  ChartAxisPatch,
  ChartFillInfo,
  ChartFillPatch,
  ChartInfo,
  ChartInsertOptions,
  ChartLabelsInfo,
  ChartLabelsPatch,
  ChartLineInfo,
  ChartLinePatch,
  ChartMarkerInfo,
  ChartMarkerPatch,
  ChartPlotInfo,
  ChartPlotPatch,
  ChartPointPatch,
  ChartPreviewOverrides,
  ChartSeriesInfo,
  ChartSeriesPatch,
  ChartData as EngineChartData,
} from '@libraz/formulon';
import {
  ChartAxisCrosses,
  ChartAxisKind,
  ChartBlanksAs,
  ChartInsertKind,
  ChartLegendPosition,
  ChartLineDash,
  ChartMarkerSymbol,
  ChartPlotType,
  ChartTitleKind,
  ChartGrouping as EngineChartGrouping,
  ChartLabelPosition as EngineChartLabelPosition,
  ChartPointState as EngineChartPointState,
  ChartSeriesDirection as EngineChartSeriesDirection,
} from '@libraz/formulon';
import { argbFromCssColor, cssColorFromArgb } from '../engine/theme-color.js';
import { autoSeriesColor } from './series-colors.js';
import type {
  ChartAxisSpec,
  ChartData,
  ChartDataLabelsSpec,
  ChartEmptyCellMode,
  ChartFillSpec,
  ChartFilterSpec,
  ChartGrouping,
  ChartKind,
  ChartLabelPosition,
  ChartLegendSpec,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartPointState,
  ChartPointStyleSpec,
  ChartScatterStyle,
  ChartSeriesDirection,
  ChartSeriesStyleSpec,
  ChartSpec,
} from './types.js';

export { argbFromCssColor, cssColorFromArgb } from '../engine/theme-color.js';

/** Why an engine chart has no `ChartSpec` projection. */
export type UnsupportedChartReason = 'chartex' | 'multi-plot' | 'plot-type';

/** Axis addressed by its role in the chart rather than by its engine id. */
export type ChartAxisRole = 'x' | 'y';

/** One engine call; the fields are the call's arguments after `(sheet, objectId)`. */
export type ChartEngineOp =
  | {
      readonly op: 'setChartType';
      readonly plot: number;
      readonly kind: ChartInsertKind;
      readonly grouping?: EngineChartGrouping;
    }
  | {
      readonly op: 'setChartSource';
      readonly source: string;
      readonly direction: EngineChartSeriesDirection;
      readonly headerRow?: boolean;
      readonly headerColumn?: boolean;
    }
  | { readonly op: 'setChartPlot'; readonly plot: number; readonly patch: ChartPlotPatch }
  | {
      readonly op: 'setChartDisplay';
      readonly blanksAs: ChartBlanksAs;
      readonly plotVisibleOnly: boolean | null;
    }
  | {
      readonly op: 'setChartTitle';
      /** Text, null for no title, or Excel's automatic title. */
      readonly title: string | null | { readonly kind: ChartTitleKind.Auto };
    }
  | {
      readonly op: 'setChartLegend';
      readonly position: ChartLegendPosition;
      readonly overlay: boolean;
    }
  | { readonly op: 'setChartAxis'; readonly axis: ChartAxisRole; readonly patch: ChartAxisPatch }
  | {
      readonly op: 'setChartSeries';
      readonly plot: number;
      readonly seriesIdx: number;
      readonly patch: ChartSeriesPatch;
    }
  | {
      readonly op: 'setChartPoint';
      readonly plot: number;
      readonly seriesIdx: number;
      readonly pointIdx: number;
      readonly patch: ChartPointPatch | null;
    }
  | {
      readonly op: 'setChartFilter';
      readonly hiddenSeries: readonly number[];
      readonly hiddenCategories: readonly number[];
    };

/** What `normalizeChartSpec` needs from the engine's split of the chart's source. */
export interface ChartNormalizeContext {
  /** Direction Excel's rules resolve `auto` to. */
  readonly direction: Exclude<ChartSeriesDirection, 'auto'>;
  /** Header flags Excel's detection resolves omitted flags to. */
  readonly headerRow: boolean;
  readonly headerColumn: boolean;
  /** Series ids in series order. */
  readonly seriesIds: readonly string[];
  /** The workbook theme's accent1..accent6 as `#rrggbb`. */
  readonly accents: readonly string[];
  /** Whether the chart has a chart style part; charts the cell inserts have none. */
  readonly styled: boolean;
  /** The kind the spec's point formats were made on, when the spec changes the kind. */
  readonly previousKind?: ChartKind;
}

interface ReadContext {
  readonly info: ChartInfo;
  readonly data: EngineChartData | null;
  readonly kind: ChartKind;
  readonly plot: ChartPlotInfo;
}

interface WriteContext {
  readonly info: ChartInfo;
  /** The spec before the change, without the fields a type change rebuilds. */
  readonly before: ChartSpec;
  readonly after: ChartSpec;
  /** Whether `setChartType` or `setChartSource` rebuilds the series. */
  readonly rebuilt: boolean;
  readonly series: readonly ChartSeriesInfo[];
  readonly ops: OpBuilder;
}

interface ChartFieldMapping<K extends keyof ChartSpec> {
  /** Kinds the field applies to. */
  readonly kinds: readonly ChartKind[];
  readonly read: (context: ReadContext) => ChartSpec[K] | undefined;
  readonly write: (context: WriteContext) => void;
}

const ALL_KINDS: readonly ChartKind[] = [
  'column',
  'bar',
  'line',
  'area',
  'pie',
  'doughnut',
  'scatter',
];
const GROUPED_KINDS: readonly ChartKind[] = ['column', 'bar', 'line', 'area'];
const BAR_KINDS: readonly ChartKind[] = ['column', 'bar'];
const CIRCULAR_KINDS: readonly ChartKind[] = ['pie', 'doughnut'];
const AXIS_KINDS: readonly ChartKind[] = ['column', 'bar', 'line', 'area', 'scatter'];
const MARKER_KINDS: readonly ChartKind[] = ['line', 'scatter'];
const SPAN_KINDS: readonly ChartKind[] = ['line', 'area', 'scatter'];

/** Fields `setChartType` rebuilds; after a type change they are written from the new spec alone. */
const TYPE_REBUILT_FIELDS: readonly (keyof ChartSpec)[] = [
  'scatterStyle',
  'gapWidth',
  'overlap',
  'firstSliceAngle',
  'holeSize',
  'palette',
  'filter',
  'dataLabels',
  'seriesStyles',
  'varyColors',
  'xAxis',
  'yAxis',
];

// ---------------------------------------------------------------------------
// Value conversions
// ---------------------------------------------------------------------------

const KIND_TO_INSERT: Readonly<Record<ChartKind, ChartInsertKind>> = {
  column: ChartInsertKind.Column,
  bar: ChartInsertKind.Bar,
  line: ChartInsertKind.Line,
  area: ChartInsertKind.Area,
  pie: ChartInsertKind.Pie,
  doughnut: ChartInsertKind.Doughnut,
  scatter: ChartInsertKind.Scatter,
};

const GROUPING_TO_ENGINE: Readonly<Record<ChartGrouping, EngineChartGrouping>> = {
  clustered: EngineChartGrouping.Clustered,
  stacked: EngineChartGrouping.Stacked,
  'percent-stacked': EngineChartGrouping.PercentStacked,
};

const DIRECTION_TO_ENGINE: Readonly<Record<ChartSeriesDirection, EngineChartSeriesDirection>> = {
  auto: EngineChartSeriesDirection.Auto,
  columns: EngineChartSeriesDirection.Columns,
  rows: EngineChartSeriesDirection.Rows,
};

const LEGEND_TO_ENGINE: Readonly<
  Record<NonNullable<ChartLegendSpec['position']>, ChartLegendPosition>
> = {
  top: ChartLegendPosition.Top,
  right: ChartLegendPosition.Right,
  bottom: ChartLegendPosition.Bottom,
  left: ChartLegendPosition.Left,
};

const DASH_TO_ENGINE: Readonly<Record<NonNullable<ChartLineSpec['dash']>, ChartLineDash>> = {
  solid: ChartLineDash.Solid,
  dash: ChartLineDash.Dash,
  dot: ChartLineDash.Dot,
  'dash-dot': ChartLineDash.DashDot,
};

const SHAPE_TO_ENGINE: Readonly<Record<NonNullable<ChartMarkerSpec['shape']>, ChartMarkerSymbol>> =
  {
    circle: ChartMarkerSymbol.Circle,
    square: ChartMarkerSymbol.Square,
    diamond: ChartMarkerSymbol.Diamond,
    triangle: ChartMarkerSymbol.Triangle,
    x: ChartMarkerSymbol.X,
    star: ChartMarkerSymbol.Star,
    plus: ChartMarkerSymbol.Plus,
    dot: ChartMarkerSymbol.Dot,
    dash: ChartMarkerSymbol.Dash,
    picture: ChartMarkerSymbol.Picture,
  };

const LABEL_POSITION_TO_ENGINE: Readonly<Record<ChartLabelPosition, EngineChartLabelPosition>> = {
  'best-fit': EngineChartLabelPosition.BestFit,
  center: EngineChartLabelPosition.Center,
  'inside-base': EngineChartLabelPosition.InsideBase,
  'inside-end': EngineChartLabelPosition.InsideEnd,
  'outside-end': EngineChartLabelPosition.OutsideEnd,
  above: EngineChartLabelPosition.Above,
  below: EngineChartLabelPosition.Below,
  left: EngineChartLabelPosition.Left,
  right: EngineChartLabelPosition.Right,
};

const BLANKS_TO_ENGINE: Readonly<Record<ChartEmptyCellMode, ChartBlanksAs>> = {
  gap: ChartBlanksAs.Gap,
  zero: ChartBlanksAs.Zero,
  span: ChartBlanksAs.Span,
};

const CROSSES_TO_ENGINE: Readonly<Record<'auto' | 'min' | 'max', ChartAxisCrosses>> = {
  auto: ChartAxisCrosses.AutoZero,
  min: ChartAxisCrosses.Min,
  max: ChartAxisCrosses.Max,
};

const invert = <K extends string, V extends number>(
  table: Readonly<Record<K, V>>,
): ReadonlyMap<V, K> =>
  new Map((Object.entries(table) as [K, V][]).map(([key, value]) => [value, key]));

const DIRECTION_FROM_ENGINE = invert(DIRECTION_TO_ENGINE);
const DASH_FROM_ENGINE = invert(DASH_TO_ENGINE);
const SHAPE_FROM_ENGINE = invert(SHAPE_TO_ENGINE);
const LABEL_POSITION_FROM_ENGINE = invert(LABEL_POSITION_TO_ENGINE);

const TICK_LABEL_POSITIONS: readonly NonNullable<ChartAxisSpec['labelPosition']>[] = [
  'nextTo',
  'low',
  'high',
  'none',
];

/** Lower-case `#rrggbb` for a colour the engine can store; other strings unchanged. */
const canonicalColor = (color: string): string => {
  const argb = argbFromCssColor(color);
  return argb === null ? color : cssColorFromArgb(argb);
};

const seriesId = (idx: number): string => `series-${idx}`;
const pointId = (idx: number): string => `point-${idx}`;
const categoryId = (kind: ChartKind, index: number): string =>
  kind === 'scatter' ? pointId(index) : `category-${index}`;

/** Trailing index of a `prefix-N` id, or null. */
const idIndex = (id: string, prefix: string): number | null => {
  if (!id.startsWith(`${prefix}-`)) return null;
  const text = id.slice(prefix.length + 1);
  return /^\d+$/.test(text) ? Number(text) : null;
};

/** The `c:idx` of the series `id` names, or null. */
export const chartSeriesIdx = (id: string): number | null => idIndex(id, 'series');

const categoryIndex = (kind: ChartKind, id: string): number | null =>
  idIndex(id, kind === 'scatter' ? 'point' : 'category');

/** `value` without its undefined entries, or undefined when nothing is left. */
const defined = <T extends object>(value: T): T | undefined => {
  const entries = Object.entries(value).filter(([, entry]) => entry !== undefined);
  return entries.length > 0 ? (Object.fromEntries(entries) as T) : undefined;
};

const nonEmpty = <T>(values: readonly T[] | undefined): readonly T[] | undefined =>
  values && values.length > 0 ? values : undefined;

const sameValue = (left: unknown, right: unknown): boolean =>
  JSON.stringify(left) === JSON.stringify(right);

// ---------------------------------------------------------------------------
// Engine → ChartSpec
// ---------------------------------------------------------------------------

const kindOfPlot = (plot: ChartPlotInfo): ChartKind | null => {
  switch (plot.type) {
    case ChartPlotType.Bar:
      return plot.barDirection === 'bar' ? 'bar' : 'column';
    case ChartPlotType.Line:
      return 'line';
    case ChartPlotType.Area:
      return 'area';
    case ChartPlotType.Pie:
      return 'pie';
    case ChartPlotType.Doughnut:
      return 'doughnut';
    case ChartPlotType.Scatter:
      return 'scatter';
    default:
      return null;
  }
};

/** Why `info` cannot be shown as a `ChartSpec`, or null when it can. */
export const unsupportedChartReason = (info: ChartInfo): UnsupportedChartReason | null => {
  if (info.isChartex) return 'chartex';
  if (info.plots.length !== 1) return 'multi-plot';
  const plot = info.plots[0];
  return plot && kindOfPlot(plot) !== null ? null : 'plot-type';
};

const readFill = (fill: ChartFillInfo): ChartFillSpec | undefined =>
  defined<ChartFillSpec>({
    color: fill.color === null ? undefined : cssColorFromArgb(fill.color),
    transparency: fill.transparency ?? undefined,
  });

// A line that is not drawn (`none`) and a marker whose symbol is None keep their
// width, dash, size and colours in the file; the spec reads them as absent.
const readLine = (line: ChartLineInfo): ChartLineSpec | undefined =>
  line.none
    ? undefined
    : defined<ChartLineSpec>({
        color: line.color === null ? undefined : cssColorFromArgb(line.color),
        transparency: line.transparency ?? undefined,
        width: line.widthPt ?? undefined,
        dash: line.dash === null ? undefined : DASH_FROM_ENGINE.get(line.dash),
      });

const readMarker = (marker: ChartMarkerInfo | null): ChartMarkerSpec | undefined => {
  if (marker === null) return undefined;
  if (marker.symbol === ChartMarkerSymbol.None) return { visible: false };
  return defined<ChartMarkerSpec>({
    visible: true,
    shape: SHAPE_FROM_ENGINE.get(marker.symbol),
    size: marker.size ?? undefined,
    fill: readFill(marker.fill),
    line: readLine(marker.line),
  });
};

const readLabels = (labels: ChartLabelsInfo | null): ChartDataLabelsSpec | undefined =>
  labels === null || labels.deleted
    ? undefined
    : defined<ChartDataLabelsSpec>({
        visible: true,
        showSeriesName: labels.showSeries,
        showCategoryName: labels.showCategory,
        showValue: labels.showValue,
        showPercentage: labels.showPercent,
        showLegendKey: labels.showLegendKey,
        separator: labels.separator ?? undefined,
        position:
          labels.position === null ? undefined : LABEL_POSITION_FROM_ENGINE.get(labels.position),
        numberFormat: labels.numberFormat ?? undefined,
        color: labels.fontColor === null ? undefined : cssColorFromArgb(labels.fontColor),
        fontSize: labels.fontSizePt ?? undefined,
      });

/** Whether every series has the same labels, and those labels. */
const sharedLabels = (
  series: readonly ChartSeriesInfo[],
): { readonly shared: boolean; readonly labels: ChartDataLabelsSpec | undefined } => {
  const each = series.map((entry) => readLabels(entry.labels));
  const first = each[0];
  return each.every((labels) => sameValue(labels, first))
    ? { shared: true, labels: first }
    : { shared: false, labels: undefined };
};

const isHorizontal = (axis: ChartAxisInfo): boolean =>
  axis.position === 'b' || axis.position === 't';

const hasRole = (kind: ChartKind, axis: ChartAxisInfo, role: ChartAxisRole): boolean => {
  if (kind === 'scatter') {
    return axis.kind === ChartAxisKind.Value && isHorizontal(axis) === (role === 'x');
  }
  return role === 'x'
    ? axis.kind === ChartAxisKind.Category || axis.kind === ChartAxisKind.Date
    : axis.kind === ChartAxisKind.Value;
};

const axisOfRole = (
  info: ChartInfo,
  plot: ChartPlotInfo,
  kind: ChartKind,
  role: ChartAxisRole,
): ChartAxisInfo | undefined =>
  info.axes.find((axis) => plot.axisIds.includes(axis.id) && hasRole(kind, axis, role));

/** Engine id of the axis playing `role` in `info`, or null when the chart has none. */
export const chartAxisId = (info: ChartInfo, role: ChartAxisRole): number | null => {
  const plot = info.plots[0];
  const kind = plot ? kindOfPlot(plot) : null;
  if (!plot || kind === null || !AXIS_KINDS.includes(kind)) return null;
  return axisOfRole(info, plot, kind, role)?.id ?? null;
};

const otherRole = (role: ChartAxisRole): ChartAxisRole => (role === 'x' ? 'y' : 'x');

// A ChartSpec axis' `crossing` is where the perpendicular axis crosses it; the engine stores
// that as `crosses` on the perpendicular axis.
const readCrossing = (cross: ChartAxisInfo | undefined): ChartAxisSpec['crossing'] => {
  switch (cross?.crosses) {
    case ChartAxisCrosses.AutoZero:
      return 'auto';
    case ChartAxisCrosses.Min:
      return 'min';
    case ChartAxisCrosses.Max:
      return 'max';
    case ChartAxisCrosses.At:
      return cross.crossesAt ?? undefined;
    default:
      return undefined;
  }
};

const readAxisTitle = (
  axis: ChartAxisInfo,
  data: EngineChartData | null,
): ChartAxisSpec['title'] => {
  switch (axis.title.kind) {
    case ChartTitleKind.Text:
      return axis.title.text;
    case ChartTitleKind.Ref:
      return data?.axisTitleTexts.find((entry) => entry.axisId === axis.id)?.text;
    case ChartTitleKind.Auto:
      return { auto: true };
    default:
      return undefined;
  }
};

const readAxis = (context: ReadContext, role: ChartAxisRole): ChartAxisSpec | undefined => {
  const { info, plot, kind } = context;
  const axis = axisOfRole(info, plot, kind, role);
  if (!axis) return undefined;
  return defined<ChartAxisSpec>({
    visible: !axis.deleted,
    logBase: axis.logBase ?? undefined,
    labelPosition: TICK_LABEL_POSITIONS.find((entry) => entry === axis.tickLabelPosition),
    reverseOrder: axis.reversed,
    crossing: readCrossing(axisOfRole(info, plot, kind, otherRole(role))),
    title: readAxisTitle(axis, context.data),
    min: axis.min ?? undefined,
    max: axis.max ?? undefined,
    majorUnit: axis.majorUnit ?? undefined,
    // A source-linked format follows the data's own format codes.
    numberFormat: axis.sourceLinked ? undefined : (axis.numberFormat ?? undefined),
    majorGridlines: axis.majorGridlines,
  });
};

const readSeriesStyles = (context: ReadContext): ChartSeriesStyleSpec[] | undefined => {
  const markers = MARKER_KINDS.includes(context.kind);
  const { shared } = sharedLabels(context.plot.series);
  const styles = context.plot.series.flatMap((series): ChartSeriesStyleSpec[] => {
    const points = series.points.flatMap((point): ChartPointStyleSpec[] => {
      const style = defined({
        fill: readFill(point.fill),
        line: readLine(point.line),
        marker: markers ? readMarker(point.marker) : undefined,
      });
      return style ? [{ pointId: categoryId(context.kind, point.idx), ...style }] : [];
    });
    const style = defined({
      // Excel keeps no fill on a line or scatter series.
      fill: markers ? undefined : readFill(series.fill),
      line: readLine(series.line),
      marker: markers ? readMarker(series.marker) : undefined,
      // An unsmoothed line reads as absent, as Excel's c:smooth val="0" and no element draw alike.
      smooth: markers && series.smooth === true ? true : undefined,
      dataLabels: shared ? undefined : readLabels(series.labels),
      points: nonEmpty(points),
    });
    return style ? [{ seriesId: seriesId(series.idx), ...style }] : [];
  });
  return styles.length > 0 ? styles : undefined;
};

const readScatterStyle = (series: readonly ChartSeriesInfo[]): ChartScatterStyle | undefined => {
  const styles = new Set(
    series.map((entry): ChartScatterStyle => {
      if (entry.line.none) return 'marker';
      // No marker element leaves Excel's automatic markers.
      const markers = entry.marker === null || entry.marker.symbol !== ChartMarkerSymbol.None;
      return markers ? 'lineMarker' : 'line';
    }),
  );
  return styles.size === 1 ? [...styles][0] : undefined;
};

const readEmptyCells = (kind: ChartKind, blanksAs: ChartBlanksAs): ChartEmptyCellMode => {
  switch (blanksAs) {
    case ChartBlanksAs.Gap:
      return 'gap';
    case ChartBlanksAs.Span:
      return emptyCellsFor(kind, 'span');
    default:
      // An absent element plots blanks as zero.
      return 'zero';
  }
};

const readGrouping = (grouping: EngineChartGrouping | null): ChartGrouping => {
  switch (grouping) {
    case EngineChartGrouping.Stacked:
      return 'stacked';
    case EngineChartGrouping.PercentStacked:
      return 'percent-stacked';
    default:
      return 'clustered';
  }
};

const readTitle = (info: ChartInfo, data: EngineChartData | null): ChartSpec['title'] => {
  switch (info.title.kind) {
    case ChartTitleKind.Text:
      return info.title.text;
    case ChartTitleKind.Ref:
      return data?.titleText ?? undefined;
    case ChartTitleKind.None:
      return '';
    case ChartTitleKind.Auto:
      return { auto: true };
    default:
      return undefined;
  }
};

const readLegend = (info: ChartInfo): ChartLegendSpec => {
  switch (info.legend.position) {
    case ChartLegendPosition.None:
      return { visible: false };
    case ChartLegendPosition.Top:
      return { visible: true, position: 'top' };
    case ChartLegendPosition.Left:
      return { visible: true, position: 'left' };
    case ChartLegendPosition.Bottom:
      return { visible: true, position: 'bottom' };
    default:
      // Top-right is drawn on the right and stays top-right until the legend changes.
      return { visible: true, position: 'right' };
  }
};

const readFilter = (context: ReadContext): ChartFilterSpec | undefined =>
  defined<ChartFilterSpec>({
    hiddenSeriesIds: nonEmpty(
      context.plot.series.filter((series) => series.filtered).map((series) => seriesId(series.idx)),
    ),
    hiddenCategoryIds: nonEmpty(
      context.info.hiddenCategories.map((index) => categoryId(context.kind, index)),
    ),
  });

// ---------------------------------------------------------------------------
// ChartSpec diff → engine operations
// ---------------------------------------------------------------------------

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

const mergeSeriesPatch = (base: ChartSeriesPatch, next: ChartSeriesPatch): ChartSeriesPatch => {
  const merged: Mutable<ChartSeriesPatch> = { ...base, ...next };
  if (base.fill && next.fill) merged.fill = { ...base.fill, ...next.fill };
  if (base.line && next.line) merged.line = { ...base.line, ...next.line };
  if (base.marker && next.marker) {
    merged.marker = {
      ...base.marker,
      ...next.marker,
      ...(base.marker.fill && next.marker.fill
        ? { fill: { ...base.marker.fill, ...next.marker.fill } }
        : {}),
      ...(base.marker.line && next.marker.line
        ? { line: { ...base.marker.line, ...next.marker.line } }
        : {}),
    };
  }
  return merged;
};

/** Collects the operations of one diff, merging those that address the same target. */
class OpBuilder {
  type: ChartEngineOp | undefined;
  source: ChartEngineOp | undefined;
  readonly plot: Mutable<ChartPlotPatch> = {};
  display: ChartEngineOp | undefined;
  title: ChartEngineOp | undefined;
  legend: ChartEngineOp | undefined;
  readonly points = new Map<string, ChartEngineOp>();
  filter: ChartEngineOp | undefined;
  private readonly axes = new Map<ChartAxisRole, ChartAxisPatch>();
  private readonly series = new Map<number, ChartSeriesPatch>();

  axis(role: ChartAxisRole, patch: ChartAxisPatch): void {
    this.axes.set(role, { ...this.axes.get(role), ...patch });
  }

  seriesPatch(idx: number, patch: ChartSeriesPatch): void {
    this.series.set(idx, mergeSeriesPatch(this.series.get(idx) ?? {}, patch));
  }

  ops(): ChartEngineOp[] {
    const ops: ChartEngineOp[] = [];
    if (this.type) ops.push(this.type);
    if (this.source) ops.push(this.source);
    if (Object.keys(this.plot).length > 0)
      ops.push({ op: 'setChartPlot', plot: 0, patch: this.plot });
    if (this.display) ops.push(this.display);
    if (this.title) ops.push(this.title);
    if (this.legend) ops.push(this.legend);
    for (const axis of ['x', 'y'] as const) {
      const patch = this.axes.get(axis);
      if (patch && Object.keys(patch).length > 0) ops.push({ op: 'setChartAxis', axis, patch });
    }
    for (const [seriesIdx, patch] of this.series) {
      if (Object.keys(patch).length > 0)
        ops.push({ op: 'setChartSeries', plot: 0, seriesIdx, patch });
    }
    ops.push(...this.points.values());
    if (this.filter) ops.push(this.filter);
    return ops;
  }
}

const fillPatch = (
  before: ChartFillSpec | undefined,
  after: ChartFillSpec | undefined,
): ChartFillPatch | undefined => {
  const patch: Mutable<ChartFillPatch> = {};
  if (before?.color !== after?.color) {
    const color = after?.color === undefined ? null : argbFromCssColor(after.color);
    // A colour the engine cannot store keeps the current one.
    if (after?.color === undefined || color !== null) patch.color = color;
  }
  if (before?.transparency !== after?.transparency) {
    patch.transparency = after?.transparency ?? null;
  }
  return Object.keys(patch).length > 0 ? patch : undefined;
};

const linePatch = (
  before: ChartLineSpec | undefined,
  after: ChartLineSpec | undefined,
): ChartLinePatch | undefined => {
  const patch: Mutable<ChartLinePatch> = { ...fillPatch(before, after) };
  if (before?.width !== after?.width) patch.widthPt = after?.width ?? null;
  if (before?.dash !== after?.dash) {
    patch.dash = after?.dash === undefined ? null : DASH_TO_ENGINE[after.dash];
  }
  return Object.keys(patch).length > 0 ? patch : undefined;
};

const markerSymbol = (
  before: ChartMarkerSpec | undefined,
  after: ChartMarkerSpec,
): ChartMarkerSymbol | undefined => {
  const wasHidden = before?.visible === false;
  if (after.visible === false) return wasHidden ? undefined : ChartMarkerSymbol.None;
  // The engine cannot write a picture marker, so a picture keeps the current symbol.
  if (after.shape !== undefined && after.shape !== 'picture') {
    return before?.shape === after.shape && !wasHidden ? undefined : SHAPE_TO_ENGINE[after.shape];
  }
  // Without a shape the automatic or unmodelled symbol stays.
  return wasHidden ? ChartMarkerSymbol.Auto : undefined;
};

const markerPatch = (
  before: ChartMarkerSpec | undefined,
  after: ChartMarkerSpec | undefined,
): ChartMarkerPatch | null | undefined => {
  if (sameValue(before, after)) return undefined;
  if (after === undefined) return null;
  const patch: Mutable<ChartMarkerPatch> = {};
  const symbol = markerSymbol(before, after);
  if (symbol !== undefined) patch.symbol = symbol;
  if (before?.size !== after.size) patch.size = after.size ?? null;
  const fill = fillPatch(before?.fill, after.fill);
  if (fill) patch.fill = fill;
  const line = linePatch(before?.line, after.line);
  if (line) patch.line = line;
  return Object.keys(patch).length > 0 ? patch : undefined;
};

/** Labels in the form the engine reads back: every flag explicit, hidden labels absent. */
const canonicalLabels = (
  labels: ChartDataLabelsSpec | undefined,
): ChartDataLabelsSpec | undefined => {
  if (!labels || labels.visible === false) return undefined;
  const flagged = [
    labels.showSeriesName,
    labels.showCategoryName,
    labels.showValue,
    labels.showPercentage,
    labels.showLegendKey,
  ].some((flag) => flag !== undefined);
  return defined<ChartDataLabelsSpec>({
    visible: true,
    showSeriesName: labels.showSeriesName ?? false,
    showCategoryName: labels.showCategoryName ?? false,
    showValue: labels.showValue ?? !flagged,
    showPercentage: labels.showPercentage ?? false,
    showLegendKey: labels.showLegendKey ?? false,
    separator: labels.separator,
    position: labels.position,
    numberFormat: labels.numberFormat,
    color: labels.color === undefined ? undefined : canonicalColor(labels.color),
    fontSize: labels.fontSize,
  });
};

const labelsPatch = (labels: ChartDataLabelsSpec | undefined): ChartLabelsPatch | null =>
  labels === undefined
    ? null
    : {
        showValue: labels.showValue === true,
        showCategory: labels.showCategoryName === true,
        showSeries: labels.showSeriesName === true,
        showPercent: labels.showPercentage === true,
        showLegendKey: labels.showLegendKey === true,
        separator: labels.separator ?? null,
        position: labels.position === undefined ? null : LABEL_POSITION_TO_ENGINE[labels.position],
        numberFormat: labels.numberFormat ?? null,
        fontColor: labels.color === undefined ? null : argbFromCssColor(labels.color),
        fontSizePt: labels.fontSize ?? null,
      };

const styleOf = (spec: ChartSpec, id: string): ChartSeriesStyleSpec | undefined =>
  spec.seriesStyles?.find((style) => style.seriesId === id);

/** A series' labels: its own, else the chart-wide ones. */
const effectiveLabels = (spec: ChartSpec, id: string): ChartDataLabelsSpec | undefined =>
  canonicalLabels(styleOf(spec, id)?.dataLabels ?? spec.dataLabels);

/** A series' fill colour: its own, else the palette entry for its position. */
const effectiveFillColor = (spec: ChartSpec, id: string, position: number): string | undefined => {
  const explicit = styleOf(spec, id)?.fill?.color;
  if (explicit !== undefined) return explicit;
  const { palette } = spec;
  return palette && palette.length > 0 ? palette[position % palette.length] : undefined;
};

/**
 * A series' fill, line and marker as the engine stores them. Bar, area and circular
 * series carry the series colour (an explicit fill colour, else the palette entry)
 * as their fill. Excel keeps no fill on a line or scatter series, so there the
 * colour goes to the line and to a shown marker's fill and outline, under any
 * colour they set themselves.
 */
const storedSeriesStyle = (
  spec: ChartSpec,
  id: string,
  position: number,
): { fill?: ChartFillSpec; line?: ChartLineSpec; marker?: ChartMarkerSpec } => {
  const style = styleOf(spec, id);
  const color = effectiveFillColor(spec, id, position);
  if (!MARKER_KINDS.includes(spec.kind))
    return { fill: { ...style?.fill, color }, line: style?.line };
  const marker = style?.marker;
  return {
    line: linelessSeries(spec) ? undefined : tinted(style?.line, color),
    marker: tintedMarker(marker, color, marker !== undefined && marker.visible !== false),
  };
};

/**
 * A point style as the engine stores it. On a line or scatter chart the point's
 * fill colour goes, like a series colour, to its line (unless the scatter style
 * draws none) and to its marker when the series or point shows one.
 */
const storedPointStyle = (
  spec: ChartSpec,
  series: ChartSeriesStyleSpec | undefined,
  point: ChartPointStyleSpec,
): ChartPointStyleSpec => {
  if (!MARKER_KINDS.includes(spec.kind)) return point;
  const { fill, ...rest } = point;
  const color = fill?.color;
  if (color === undefined) return rest;
  const shown =
    (point.marker ?? series?.marker) !== undefined &&
    (point.marker?.visible ?? series?.marker?.visible) !== false;
  const line = linelessSeries(spec) ? point.line : tinted(point.line, color);
  const marker = tintedMarker(point.marker, color, shown);
  return { ...rest, ...(line ? { line } : {}), ...(marker ? { marker } : {}) };
};

/** `own` with `color` where it sets no colour of its own. */
const tinted = <T extends { readonly color?: string }>(
  own: T | undefined,
  color: string | undefined,
): T | undefined =>
  color === undefined || own?.color !== undefined ? own : ({ ...own, color } as T);

const tintedMarker = (
  marker: ChartMarkerSpec | undefined,
  color: string | undefined,
  shown: boolean,
): ChartMarkerSpec | undefined =>
  !shown || color === undefined
    ? marker
    : { ...marker, fill: tinted(marker?.fill, color), line: tinted(marker?.line, color) };

/** `value` without its colour, or undefined when nothing else is left. */
const uncolored = <T extends { readonly color?: string }>(value: T | undefined): T | undefined => {
  if (value?.color === undefined) return value;
  const { color: _color, ...rest } = value;
  return Object.keys(rest).length > 0 ? (rest as T) : undefined;
};

/**
 * `style` without the colour a palette supplies: its fill colour, and on a line or
 * scatter chart, where the series colour is stored, its line and marker colours.
 */
export const withoutSeriesColor = (
  kind: ChartKind,
  style: ChartSeriesStyleSpec,
): ChartSeriesStyleSpec => {
  const markers = MARKER_KINDS.includes(kind);
  const marker =
    markers && style.marker
      ? defined<ChartMarkerSpec>({
          ...style.marker,
          fill: uncolored(style.marker.fill),
          line: uncolored(style.marker.line),
        })
      : style.marker;
  const stripped = {
    ...style,
    fill: uncolored(style.fill),
    line: markers ? uncolored(style.line) : style.line,
    marker: style.marker ? (marker ?? {}) : undefined,
  };
  return Object.fromEntries(
    Object.entries(stripped).filter(([, value]) => value !== undefined),
  ) as unknown as ChartSeriesStyleSpec;
};

/**
 * The colour a series shows: its fill colour, or on a line or scatter chart its
 * line colour (its marker fill when the scatter style draws no line).
 */
export const seriesColorOf = (
  spec: Pick<ChartSpec, 'kind' | 'scatterStyle'>,
  style: ChartSeriesStyleSpec | undefined,
): string | undefined =>
  MARKER_KINDS.includes(spec.kind)
    ? ((linelessSeries(spec) ? undefined : style?.line?.color) ??
      style?.marker?.fill?.color ??
      style?.fill?.color)
    : style?.fill?.color;

const writeFillColors = (context: WriteContext): void => {
  // A line or scatter series' colour is written with its line and markers.
  if (MARKER_KINDS.includes(context.after.kind)) return;
  context.series.forEach((series, position) => {
    const id = seriesId(series.idx);
    const fill = fillPatch(
      { color: effectiveFillColor(context.before, id, position) },
      { color: effectiveFillColor(context.after, id, position) },
    );
    if (fill) context.ops.seriesPatch(series.idx, { fill });
  });
};

const writeLabels = (context: WriteContext): void => {
  for (const series of context.series) {
    const id = seriesId(series.idx);
    const after = effectiveLabels(context.after, id);
    if (!sameValue(effectiveLabels(context.before, id), after)) {
      context.ops.seriesPatch(series.idx, { labels: labelsPatch(after) });
    }
  }
};

const writePoints = (
  context: WriteContext,
  seriesIdx: number,
  before: readonly ChartPointStyleSpec[] | undefined,
  after: readonly ChartPointStyleSpec[] | undefined,
  markers: boolean,
): void => {
  const ids = new Set([...(before ?? []), ...(after ?? [])].map((point) => point.pointId));
  for (const id of ids) {
    const pointIdx = categoryIndex(context.after.kind, id);
    if (pointIdx === null) continue;
    const from = before?.find((point) => point.pointId === id);
    const to = after?.find((point) => point.pointId === id);
    let patch: ChartPointPatch | null = null;
    if (to) {
      const next: Mutable<ChartPointPatch> = {};
      const fill = fillPatch(from?.fill, to.fill);
      if (fill) next.fill = fill;
      const line = linePatch(from?.line, to.line);
      if (line) next.line = line;
      const marker = markers ? markerPatch(from?.marker, to.marker) : undefined;
      if (marker !== undefined) next.marker = marker;
      if (Object.keys(next).length === 0) continue;
      patch = next;
    }
    context.ops.points.set(`${seriesIdx}:${pointIdx}`, {
      op: 'setChartPoint',
      plot: 0,
      seriesIdx,
      pointIdx,
      patch,
    });
  }
};

const writeSeriesStyles = (context: WriteContext): void => {
  const markers = MARKER_KINDS.includes(context.after.kind);
  context.series.forEach((series, position) => {
    const id = seriesId(series.idx);
    const before = styleOf(context.before, id);
    const after = styleOf(context.after, id);
    const patch: Mutable<ChartSeriesPatch> = {};
    // A line or scatter series has no fill; its colour comes with the line and markers.
    const from = markers ? storedSeriesStyle(context.before, id, position) : before;
    const to = markers ? storedSeriesStyle(context.after, id, position) : after;
    // The colour is the `palette` entry's: an explicit colour overrides the palette.
    const fill = markers
      ? undefined
      : fillPatch(
          { transparency: before?.fill?.transparency },
          { transparency: after?.fill?.transparency },
        );
    if (fill) patch.fill = fill;
    // The scatter style owns the line of a marker-only series.
    const line = linelessSeries(context.after) ? undefined : linePatch(from?.line, to?.line);
    if (line) patch.line = line;
    const marker = markers ? markerPatch(from?.marker, to?.marker) : undefined;
    if (marker !== undefined) patch.marker = marker;
    if (markers && (before?.smooth === true) !== (after?.smooth === true)) {
      patch.smooth = after?.smooth === true;
    }
    if (Object.keys(patch).length > 0) context.ops.seriesPatch(series.idx, patch);
    writePoints(
      context,
      series.idx,
      before?.points?.map((point) => storedPointStyle(context.before, before, point)),
      after?.points?.map((point) => storedPointStyle(context.after, after, point)),
      markers,
    );
  });
};

const writeScatterStyle = (context: WriteContext): void => {
  const before = context.before.scatterStyle;
  const after = context.after.scatterStyle;
  if (after === undefined || after === before) return;
  const lines = (style: ChartScatterStyle | undefined): boolean => style !== 'marker';
  const markers = (style: ChartScatterStyle | undefined): boolean => style !== 'line';
  for (const series of context.series) {
    const patch: Mutable<ChartSeriesPatch> = {};
    if (before === undefined || lines(before) !== lines(after)) {
      patch.line = { none: !lines(after) };
    }
    if (before === undefined || markers(before) !== markers(after)) {
      patch.marker = { symbol: markers(after) ? ChartMarkerSymbol.Auto : ChartMarkerSymbol.None };
    }
    context.ops.seriesPatch(series.idx, patch);
  }
};

const engineGrouping = (spec: ChartSpec): EngineChartGrouping => {
  const grouping = spec.grouping ?? 'clustered';
  return grouping === 'clustered' && (spec.kind === 'line' || spec.kind === 'area')
    ? EngineChartGrouping.Standard
    : GROUPING_TO_ENGINE[grouping];
};

const typeChanged = (before: ChartSpec, after: ChartSpec): boolean =>
  before.kind !== after.kind ||
  (GROUPED_KINDS.includes(after.kind) &&
    (before.grouping ?? 'clustered') !== (after.grouping ?? 'clustered'));

const sourceChanged = (before: ChartSpec, after: ChartSpec): boolean =>
  before.seriesDirection !== after.seriesDirection ||
  before.headerRow !== after.headerRow ||
  before.headerColumn !== after.headerColumn;

const writeType = (context: WriteContext): void => {
  const { after } = context;
  if (!typeChanged(context.before, after)) return;
  context.ops.type = {
    op: 'setChartType',
    plot: 0,
    kind: KIND_TO_INSERT[after.kind],
    ...(GROUPED_KINDS.includes(after.kind) ? { grouping: engineGrouping(after) } : {}),
  };
};

const writeSource = (context: WriteContext): void => {
  const { source } = context.info;
  const { after } = context;
  if (!source || !sourceChanged(context.before, after)) return;
  context.ops.source = {
    op: 'setChartSource',
    source: source.range,
    direction: DIRECTION_TO_ENGINE[after.seriesDirection ?? 'auto'],
    ...(after.headerRow === undefined ? {} : { headerRow: after.headerRow }),
    ...(after.headerColumn === undefined ? {} : { headerColumn: after.headerColumn }),
  };
};

const writePlotField =
  (field: keyof ChartPlotPatch) =>
  (context: WriteContext): void => {
    const after = context.after[field];
    if (after === undefined || after === context.before[field]) return;
    (context.ops.plot as Record<keyof ChartPlotPatch, number | boolean>)[field] = after;
  };

const emptyCellsFor = (kind: ChartKind, mode: ChartEmptyCellMode): ChartEmptyCellMode =>
  mode === 'span' && !SPAN_KINDS.includes(kind) ? 'gap' : mode;

const writeEmptyCells = (context: WriteContext): void => {
  const { kind, emptyCells: after } = context.after;
  const before = context.before.emptyCells;
  if (after === undefined) return;
  if (before !== undefined && emptyCellsFor(kind, before) === emptyCellsFor(kind, after)) return;
  context.ops.display = {
    op: 'setChartDisplay',
    blanksAs: BLANKS_TO_ENGINE[emptyCellsFor(kind, after)],
    plotVisibleOnly: context.info.plotVisibleOnly,
  };
};

const writeTitle = (context: WriteContext): void => {
  const after = context.after.title;
  // An omitted title keeps the current one.
  if (after === undefined || sameValue(after, context.before.title)) return;
  context.ops.title = {
    op: 'setChartTitle',
    title: typeof after !== 'string' ? { kind: ChartTitleKind.Auto } : after === '' ? null : after,
  };
};

/** A shown legend without a position sits at the bottom, where `insertChart` puts it. */
const canonicalLegend = (legend: ChartLegendSpec | undefined): ChartLegendSpec | undefined => {
  if (!legend) return undefined;
  if (legend.visible === false) return { visible: false };
  return { visible: true, position: legend.position ?? 'bottom' };
};

const writeLegend = (context: WriteContext): void => {
  const after = canonicalLegend(context.after.legend);
  if (after === undefined || sameValue(canonicalLegend(context.before.legend), after)) return;
  context.ops.legend = {
    op: 'setChartLegend',
    position: after.position ? LEGEND_TO_ENGINE[after.position] : ChartLegendPosition.None,
    overlay: context.info.legend.overlay,
  };
};

const filterIndices = (
  spec: ChartSpec,
): { readonly series: number[]; readonly categories: number[] } => {
  const indices = (ids: readonly string[] | undefined, parse: (id: string) => number | null) =>
    [...new Set((ids ?? []).map(parse).filter((index): index is number => index !== null))].sort(
      (left, right) => left - right,
    );
  return {
    series: indices(spec.filter?.hiddenSeriesIds, chartSeriesIdx),
    categories: indices(spec.filter?.hiddenCategoryIds, (id) => categoryIndex(spec.kind, id)),
  };
};

const writeFilter = (context: WriteContext): void => {
  const before = filterIndices({ ...context.before, kind: context.after.kind });
  const after = filterIndices(context.after);
  const empty = (indices: typeof after): boolean =>
    indices.series.length === 0 && indices.categories.length === 0;
  // A rebuild clears the engine's filter, so a filter on either side is written again.
  if (context.rebuilt ? empty(before) && empty(after) : sameValue(before, after)) return;
  context.ops.filter = {
    op: 'setChartFilter',
    hiddenSeries: after.series,
    hiddenCategories: after.categories,
  };
};

const writeAxis =
  (role: ChartAxisRole) =>
  (context: WriteContext): void => {
    const key = role === 'x' ? 'xAxis' : 'yAxis';
    const before = context.before[key];
    const after = context.after[key];
    if (sameValue(before, after)) return;
    const changed = (field: keyof ChartAxisSpec): boolean =>
      !sameValue(before?.[field], after?.[field]);
    const patch: Mutable<ChartAxisPatch> = {};
    if (changed('visible')) patch.visible = after?.visible ?? true;
    if (changed('title'))
      patch.title =
        typeof after?.title === 'object'
          ? { kind: ChartTitleKind.Auto }
          : after?.title
            ? after.title
            : null;
    if (changed('min')) patch.min = after?.min ?? null;
    if (changed('max')) patch.max = after?.max ?? null;
    if (changed('majorUnit')) patch.majorUnit = after?.majorUnit ?? null;
    if (changed('numberFormat')) patch.numberFormat = after?.numberFormat ?? null;
    if (changed('majorGridlines')) patch.majorGridlines = after?.majorGridlines ?? true;
    if (changed('reverseOrder')) patch.reversed = after?.reverseOrder ?? false;
    if (changed('logBase')) patch.logBase = after?.logBase ?? null;
    if (changed('labelPosition')) patch.tickLabelPosition = after?.labelPosition ?? null;
    context.ops.axis(role, patch);
    if (changed('crossing')) {
      const crossing = after?.crossing ?? 'auto';
      context.ops.axis(
        otherRole(role),
        typeof crossing === 'number'
          ? { crosses: ChartAxisCrosses.At, crossesAt: crossing }
          : { crosses: CROSSES_TO_ENGINE[crossing] },
      );
    }
  };

// ---------------------------------------------------------------------------
// The mapping table
// ---------------------------------------------------------------------------

/**
 * One entry per `ChartSpec` field. Keyed by `keyof ChartSpec`, so adding or
 * removing a field fails type checking until the table follows.
 */
export const CHART_SPEC_FIELDS: { readonly [K in keyof ChartSpec]-?: ChartFieldMapping<K> } = {
  kind: { kinds: ALL_KINDS, read: (context) => context.kind, write: writeType },
  grouping: {
    kinds: GROUPED_KINDS,
    read: (context) => readGrouping(context.plot.grouping),
    write: writeType,
  },
  scatterStyle: {
    kinds: ['scatter'],
    read: (context) => readScatterStyle(context.plot.series),
    write: writeScatterStyle,
  },
  seriesDirection: {
    kinds: ALL_KINDS,
    read: (context) =>
      context.info.source ? DIRECTION_FROM_ENGINE.get(context.info.source.direction) : undefined,
    write: writeSource,
  },
  emptyCells: {
    kinds: ALL_KINDS,
    read: (context) => readEmptyCells(context.kind, context.info.blanksAs),
    write: writeEmptyCells,
  },
  gapWidth: {
    kinds: BAR_KINDS,
    read: (context) => context.plot.gapWidth ?? undefined,
    write: writePlotField('gapWidth'),
  },
  overlap: {
    kinds: BAR_KINDS,
    read: (context) => context.plot.overlap ?? undefined,
    write: writePlotField('overlap'),
  },
  firstSliceAngle: {
    kinds: CIRCULAR_KINDS,
    read: (context) => context.plot.firstSliceAngle ?? undefined,
    write: writePlotField('firstSliceAngle'),
  },
  holeSize: {
    kinds: ['doughnut'],
    read: (context) => context.plot.holeSize ?? undefined,
    write: writePlotField('holeSize'),
  },
  headerRow: {
    kinds: ALL_KINDS,
    read: (context) => context.info.source?.headerRow,
    write: writeSource,
  },
  headerColumn: {
    kinds: ALL_KINDS,
    read: (context) => context.info.source?.headerColumn,
    write: writeSource,
  },
  title: {
    kinds: ALL_KINDS,
    read: (context) => readTitle(context.info, context.data),
    write: writeTitle,
  },
  // The engine has no palette: each series carries its own colour.
  palette: { kinds: ALL_KINDS, read: () => undefined, write: writeFillColors },
  legend: { kinds: ALL_KINDS, read: (context) => readLegend(context.info), write: writeLegend },
  filter: { kinds: ALL_KINDS, read: readFilter, write: writeFilter },
  // Shared labels read here; labels that differ by series read into `seriesStyles`.
  dataLabels: {
    kinds: ALL_KINDS,
    read: (context) => sharedLabels(context.plot.series).labels,
    write: writeLabels,
  },
  seriesStyles: { kinds: ALL_KINDS, read: readSeriesStyles, write: writeSeriesStyles },
  varyColors: {
    kinds: ALL_KINDS,
    read: (context) => context.plot.varyColors,
    write: writePlotField('varyColors'),
  },
  xAxis: { kinds: AXIS_KINDS, read: (context) => readAxis(context, 'x'), write: writeAxis('x') },
  yAxis: { kinds: AXIS_KINDS, read: (context) => readAxis(context, 'y'), write: writeAxis('y') },
};

const FIELD_KEYS = Object.keys(CHART_SPEC_FIELDS) as (keyof ChartSpec)[];

const appliesTo = (key: keyof ChartSpec, kind: ChartKind): boolean =>
  CHART_SPEC_FIELDS[key].kinds.includes(kind);

/**
 * Projects an engine chart onto a `ChartSpec`; null when `unsupportedChartReason`
 * names a reason. `data` supplies the text of reference titles.
 */
export const chartSpecFromEngine = (
  info: ChartInfo,
  data: EngineChartData | null = null,
): ChartSpec | null => {
  const plot = info.plots[0];
  const kind = plot ? kindOfPlot(plot) : null;
  if (!plot || kind === null || unsupportedChartReason(info) !== null) return null;
  const context: ReadContext = { info, data, kind, plot };
  const spec: Record<string, unknown> = {};
  for (const key of FIELD_KEYS) {
    if (!appliesTo(key, kind)) continue;
    const value = CHART_SPEC_FIELDS[key].read(context);
    if (value !== undefined) spec[key] = value;
  }
  return spec as unknown as ChartSpec;
};

/** The `insertChart` options that create `spec`'s plot, split and legend. */
export const chartInsertOptions = (
  spec: ChartSpec,
): Required<Pick<ChartInsertOptions, 'kind' | 'direction'>> &
  Pick<ChartInsertOptions, 'grouping' | 'headerRow' | 'headerColumn' | 'legend'> => {
  const legend = canonicalLegend(spec.legend);
  return {
    kind: KIND_TO_INSERT[spec.kind],
    ...(GROUPED_KINDS.includes(spec.kind) ? { grouping: engineGrouping(spec) } : {}),
    direction: DIRECTION_TO_ENGINE[spec.seriesDirection ?? 'auto'],
    ...(spec.headerRow === undefined ? {} : { headerRow: spec.headerRow }),
    ...(spec.headerColumn === undefined ? {} : { headerColumn: spec.headerColumn }),
    ...(legend === undefined
      ? {}
      : {
          legend: legend.position ? LEGEND_TO_ENGINE[legend.position] : ChartLegendPosition.None,
        }),
  };
};

/** The `previewChartData` overrides that resolve the series as `spec` splits and plots them. */
export const chartPreviewOverrides = (spec: ChartSpec): ChartPreviewOverrides => ({
  direction: DIRECTION_TO_ENGINE[spec.seriesDirection ?? 'auto'],
  ...(spec.headerRow === undefined ? {} : { headerRow: spec.headerRow }),
  ...(spec.headerColumn === undefined ? {} : { headerColumn: spec.headerColumn }),
  ...(spec.emptyCells === undefined
    ? {}
    : { blanksAs: BLANKS_TO_ENGINE[emptyCellsFor(spec.kind, spec.emptyCells)] }),
});

/**
 * The engine operations that turn the chart `info`, projected as `before`, into
 * `after`, in execution order. Series operations address the series of `info`;
 * `setChartSource` rebuilds them, so changes aimed at the rebuilt series need a
 * second diff against the re-read chart.
 */
export const chartSpecOps = (
  info: ChartInfo,
  before: ChartSpec,
  after: ChartSpec,
): ChartEngineOp[] => {
  const retyped = typeChanged(before, after);
  const baseline: Record<string, unknown> = { ...before };
  if (retyped) for (const key of TYPE_REBUILT_FIELDS) delete baseline[key];
  const context: WriteContext = {
    info,
    before: baseline as unknown as ChartSpec,
    after,
    rebuilt: retyped || sourceChanged(before, after),
    series: info.plots[0]?.series ?? [],
    ops: new OpBuilder(),
  };
  const written = new Set<(context: WriteContext) => void>();
  for (const key of FIELD_KEYS) {
    const { write } = CHART_SPEC_FIELDS[key];
    if (!appliesTo(key, after.kind) || written.has(write)) continue;
    written.add(write);
    write(context);
  }
  return context.ops.ops();
};

// ---------------------------------------------------------------------------
// Engine chart data → drawing data
// ---------------------------------------------------------------------------

const POINT_STATES: Readonly<Record<number, ChartPointState>> = {
  [EngineChartPointState.Value]: 'value',
  [EngineChartPointState.Gap]: 'gap',
  [EngineChartPointState.Connect]: 'connect',
};

const pointState = (state: number | undefined): ChartPointState =>
  (state === undefined ? undefined : POINT_STATES[state]) ?? 'gap';

const finite = (value: number | undefined): number | null =>
  value !== undefined && Number.isFinite(value) ? value : null;

/** Converts the engine's resolved series into drawing data with the adapter's ids. */
export const chartDataFromEngine = (data: EngineChartData): ChartData => {
  if (data.domain === 'xy') {
    return {
      domain: 'xy',
      series: data.series.map((series) => ({
        id: seriesId(series.idx),
        ...(series.name === null ? {} : { name: series.name }),
        formatCode: series.formatCode,
        xFormatCode: series.xFormatCode,
        points: Array.from({ length: data.pointCount }, (_, index) => {
          const state = pointState(series.states[index]);
          return {
            id: pointId(index),
            // Without X values Excel plots the points at 1, 2, 3, ...
            x: series.x ? finite(series.x[index]) : index + 1,
            y: state === 'value' ? finite(series.values[index]) : null,
            state,
          };
        }),
      })),
    };
  }
  return {
    domain: 'category',
    categories: data.categories.labels,
    categoryIds: data.categories.labels.map((_label, index) => `category-${index}`),
    categoryFormatCode: data.categories.formatCode,
    series: data.series.map((series) => {
      const states = Array.from({ length: data.pointCount }, (_, index) =>
        pointState(series.states[index]),
      );
      return {
        id: seriesId(series.idx),
        ...(series.name === null ? {} : { name: series.name }),
        values: states.map((state, index) =>
          state === 'value' ? finite(series.values[index]) : null,
        ),
        states,
        formatCode: series.formatCode,
      };
    }),
  };
};

// ---------------------------------------------------------------------------
// Round-trip equivalence
// ---------------------------------------------------------------------------

const canonicalFill = (fill: ChartFillSpec | undefined): ChartFillSpec | undefined =>
  fill &&
  defined<ChartFillSpec>({
    color: fill.color === undefined ? undefined : canonicalColor(fill.color),
    transparency: fill.transparency,
  });

const canonicalLine = (line: ChartLineSpec | undefined): ChartLineSpec | undefined =>
  line &&
  defined<ChartLineSpec>({
    color: line.color === undefined ? undefined : canonicalColor(line.color),
    transparency: line.transparency,
    width: line.width,
    dash: line.dash,
  });

const canonicalMarker = (marker: ChartMarkerSpec | undefined): ChartMarkerSpec | undefined => {
  if (!marker) return undefined;
  if (marker.visible === false) return { visible: false };
  return defined<ChartMarkerSpec>({
    visible: true,
    shape: marker.shape,
    size: marker.size,
    fill: canonicalFill(marker.fill),
    line: canonicalLine(marker.line),
  });
};

/** A marker-only scatter series draws no line; a scatter chart without a style is marker-only, as cell draws it. */
const linelessSeries = (spec: Pick<ChartSpec, 'kind' | 'scatterStyle'>): boolean =>
  spec.kind === 'scatter' && (spec.scatterStyle ?? 'marker') === 'marker';

/** Kinds where a new data point copies the series fill and outline before its own patch. */
const FILLED_POINT_KINDS: readonly ChartKind[] = ['column', 'bar', 'area'];

const over = <T extends object>(base: T | undefined, own: T | undefined): T | undefined =>
  base || own ? ({ ...base, ...own } as T) : undefined;

const overMarker = (
  base: ChartMarkerSpec | undefined,
  own: ChartMarkerSpec | undefined,
): ChartMarkerSpec | undefined => {
  const merged = over(base, own);
  if (!merged) return undefined;
  const fill = over(base?.fill, own?.fill);
  const line = over(base?.line, own?.line);
  return { ...merged, ...(fill ? { fill } : {}), ...(line ? { line } : {}) };
};

/**
 * A point style as the engine stores it. A patch lands on the point's existing
 * c:dPt when there is one, else on a new c:dPt that starts as Excel's measured
 * copy of the series: bar, column and area copy the fill and line; line takes no
 * fill and copies the line and marker; scatter takes no fill and copies line and
 * marker only when the patch sets one of them. A bar or column chart with vary
 * colours already has a c:dPt per point in the accent form (an accent fill and
 * no line), which the patch replaces rather than a copy of the series.
 */
const normalizePoint = (
  spec: ChartSpec,
  point: ChartPointStyleSpec,
  series: { fill?: ChartFillSpec; line?: ChartLineSpec; marker?: ChartMarkerSpec },
  retyped: boolean,
): Omit<ChartPointStyleSpec, 'pointId'> | undefined => {
  const { kind } = spec;
  const markers = MARKER_KINDS.includes(kind);
  // A type change converts every formatted point in place, so the patch lands on an existing c:dPt.
  const existingPoint =
    retyped || ((kind === 'bar' || kind === 'column') && spec.varyColors === true);
  const copiesSeries =
    !existingPoint &&
    (FILLED_POINT_KINDS.includes(kind) ||
      kind === 'line' ||
      (kind === 'scatter' && (point.line !== undefined || point.marker !== undefined)));
  const fill = markers
    ? undefined
    : copiesSeries && FILLED_POINT_KINDS.includes(kind)
      ? over(series.fill, point.fill)
      : point.fill;
  return defined({
    fill: canonicalFill(fill),
    line: canonicalLine(copiesSeries ? over(series.line, point.line) : point.line),
    marker: markers
      ? canonicalMarker(copiesSeries ? overMarker(series.marker, point.marker) : point.marker)
      : undefined,
  });
};

/** A marker symbol as the engine writes it. */
type MarkerSymbolName = 'none' | NonNullable<ChartMarkerSpec['shape']>;

/** The symbols an automatic series marker takes, by series idx. */
const AUTO_SYMBOLS: readonly MarkerSymbolName[] = [
  'diamond',
  'square',
  'triangle',
  'x',
  'star',
  'circle',
  'plus',
  'dot',
  'dash',
];

/** The symbol a marker spec names: none when hidden, undefined when automatic. */
const markerSymbolOf = (marker: ChartMarkerSpec | undefined): MarkerSymbolName | undefined =>
  marker?.visible === false ? 'none' : marker?.shape;

/**
 * A written marker as the engine completes it, the form Excel re-saves unchanged:
 * a hidden marker keeps only its symbol; a shown one takes size 5 when it has none
 * and, in a styled chart, a fill and a 0.75 pt outline in the series colour where
 * unset. `symbol` is the concrete symbol.
 */
const settledMarker = (
  marker: ChartMarkerSpec,
  symbol: MarkerSymbolName,
  styled: boolean,
  color: string | undefined,
): ChartMarkerSpec => {
  if (symbol === 'none') return { visible: false };
  const sized: ChartMarkerSpec = {
    ...marker,
    visible: true,
    shape: symbol,
    size: marker.size ?? 5,
  };
  if (!styled) return sized;
  return {
    ...sized,
    fill: marker.fill?.color === undefined ? { ...marker.fill, color } : marker.fill,
    line: {
      ...marker.line,
      width: marker.line?.width ?? 0.75,
      color: marker.line?.color ?? color,
    },
  };
};

const normalizeSeriesStyles = (
  spec: ChartSpec,
  seriesIds: readonly string[],
  labelsShared: boolean,
  context: ChartNormalizeContext,
): ChartSeriesStyleSpec[] | undefined => {
  const markers = MARKER_KINDS.includes(spec.kind);
  const retyped = context.previousKind !== undefined && context.previousKind !== spec.kind;
  const styles = seriesIds.flatMap((id, position): ChartSeriesStyleSpec[] => {
    const style = styleOf(spec, id);
    const idx = chartSeriesIdx(id) ?? position;
    const color = context.styled ? autoSeriesColor(context.accents, idx) : undefined;
    const unsettled = storedSeriesStyle(spec, id, position);
    const seriesSymbol =
      markerSymbolOf(unsettled.marker) ?? AUTO_SYMBOLS[idx % AUTO_SYMBOLS.length] ?? 'none';
    const stored = {
      ...unsettled,
      marker:
        markers && unsettled.marker
          ? settledMarker(unsettled.marker, seriesSymbol, context.styled, color)
          : unsettled.marker,
    };
    const points = (style?.points ?? []).flatMap((point): ChartPointStyleSpec[] => {
      const own = storedPointStyle(spec, style, point);
      const normalized = normalizePoint(
        spec,
        own,
        {
          ...(stored.fill ? { fill: stored.fill } : {}),
          ...(stored.line ? { line: stored.line } : {}),
          ...(stored.marker ? { marker: stored.marker } : {}),
        },
        retyped,
      );
      // A point marker the point writes is completed too; an automatic one takes the series' symbol.
      const marker =
        markers && own.marker
          ? canonicalMarker(
              settledMarker(
                normalized?.marker ?? own.marker,
                markerSymbolOf(own.marker) ?? seriesSymbol,
                context.styled,
                color,
              ),
            )
          : normalized?.marker;
      const settled = defined({ ...normalized, marker });
      return settled ? [{ pointId: point.pointId, ...settled }] : [];
    });
    const normalized = defined({
      fill: canonicalFill(stored.fill),
      line: canonicalLine(stored.line),
      marker: markers ? canonicalMarker(stored.marker) : undefined,
      smooth: markers && style?.smooth === true ? true : undefined,
      dataLabels: labelsShared ? undefined : effectiveLabels(spec, id),
      points: nonEmpty(points),
    });
    return normalized ? [{ seriesId: id, ...normalized }] : [];
  });
  return nonEmpty(styles) as ChartSeriesStyleSpec[] | undefined;
};

const normalizeAxis = (axis: ChartAxisSpec | undefined): ChartAxisSpec | undefined =>
  axis && defined<ChartAxisSpec>({ ...axis, title: axis.title === '' ? undefined : axis.title });

const normalizeFilter = (
  spec: ChartSpec,
  seriesIds: readonly string[],
): ChartFilterSpec | undefined => {
  const hidden = new Set(spec.filter?.hiddenSeriesIds ?? []);
  return defined<ChartFilterSpec>({
    hiddenSeriesIds: nonEmpty(seriesIds.filter((id) => hidden.has(id))),
    hiddenCategoryIds: nonEmpty(
      filterIndices(spec).categories.map((index) => categoryId(spec.kind, index)),
    ),
  });
};

/**
 * Brings `spec` to the form the engine reads back after creating a chart from it:
 * the palette becomes each series' fill colour (on a line or scatter chart, its
 * line and marker colour, with no series fill), an omitted title is automatic,
 * `auto` and omitted header flags
 * take the engine's split in `context`, omitted `emptyCells` / `grouping` take the
 * values the engine writes, labels, legend, filter and colours take their stored
 * form, and fields outside the kind are dropped.
 */
export const normalizeChartSpec = (spec: ChartSpec, context: ChartNormalizeContext): ChartSpec => {
  const { kind } = spec;
  const { seriesIds } = context;
  const labels = seriesIds.map((id) => effectiveLabels(spec, id));
  const labelsShared = labels.every((entry) => sameValue(entry, labels[0]));
  const normalized: Record<string, unknown> = {
    ...spec,
    grouping: spec.grouping ?? 'clustered',
    seriesDirection:
      spec.seriesDirection === undefined || spec.seriesDirection === 'auto'
        ? context.direction
        : spec.seriesDirection,
    headerRow: spec.headerRow ?? context.headerRow,
    headerColumn: spec.headerColumn ?? context.headerColumn,
    emptyCells: emptyCellsFor(kind, spec.emptyCells ?? 'gap'),
    // A chart created without a title text keeps the template's automatic title.
    title: spec.title ?? { auto: true },
    palette: undefined,
    legend: canonicalLegend(spec.legend),
    filter: normalizeFilter(spec, seriesIds),
    dataLabels: labelsShared ? labels[0] : undefined,
    seriesStyles: normalizeSeriesStyles(spec, seriesIds, labelsShared, context),
    xAxis: normalizeAxis(spec.xAxis),
    yAxis: normalizeAxis(spec.yAxis),
  };
  for (const key of FIELD_KEYS) {
    if (!appliesTo(key, kind) || normalized[key] === undefined) delete normalized[key];
  }
  return normalized as unknown as ChartSpec;
};
