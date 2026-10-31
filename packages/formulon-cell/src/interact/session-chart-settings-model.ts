/** Draft model for the chart settings dialog: shapes, cloning, bounds and patch building. */

import { seriesColorOf, withoutSeriesColor } from '../charts/engine-adapter.js';
import { prepareChartData } from '../charts/filters.js';
import { DEFAULT_THEME_ACCENTS } from '../charts/series-colors.js';
import type {
  ChartAxisSpec,
  ChartData,
  ChartDataLabelSpec,
  ChartEmptyCellMode,
  ChartFillSpec,
  ChartGrouping,
  ChartKind,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartPointStyleSpec,
  ChartScatterStyle,
  ChartSeriesDirection,
  ChartSeriesStyleSpec,
  ChartSpec,
} from '../charts/types.js';
import type { SessionChart } from '../store/store.js';
import {
  DEFAULT_LABELS,
  type SessionChartSettingsLabels,
} from './session-chart-settings-labels.js';

/** A chart the settings surface can edit: one the renderer can draw. */
export type SessionChartSettingsShape = SessionChart & { kind: ChartKind };
export type SessionChartSettingsPatch = {
  -readonly [Key in keyof ChartSpec]?: ChartSpec[Key];
};

/** Cell references of a series to add, as typed; `values` are a scatter chart's Y values. */
export interface SessionChartSeriesRefs {
  name?: string;
  values: string;
  categories?: string;
  x?: string;
}

const cloneFill = (fill?: ChartFillSpec): ChartFillSpec | undefined =>
  fill ? { ...fill } : undefined;

const cloneLine = (line?: ChartLineSpec): ChartLineSpec | undefined =>
  line ? { ...line } : undefined;

const cloneMarker = (marker?: ChartMarkerSpec): ChartMarkerSpec | undefined =>
  marker
    ? {
        ...marker,
        fill: cloneFill(marker.fill),
        line: cloneLine(marker.line),
      }
    : undefined;

const cloneStyles = (
  styles?: readonly ChartSeriesStyleSpec[],
): ChartSeriesStyleSpec[] | undefined =>
  styles?.map((style) => ({
    ...style,
    fill: cloneFill(style.fill),
    line: cloneLine(style.line),
    marker: cloneMarker(style.marker),
    ...(style.points
      ? {
          points: style.points.map((point) => ({
            ...point,
            fill: cloneFill(point.fill),
            line: cloneLine(point.line),
            marker: cloneMarker(point.marker),
          })),
        }
      : {}),
  }));

export interface ChartDraft {
  kind: SessionChartSettingsShape['kind'];
  kindBaseline: SessionChartSettingsShape['kind'];
  kindDirty: boolean;
  kindChanged: boolean;
  scatterStyle: ChartScatterStyle;
  scatterStyleBaseline: ChartScatterStyle;
  scatterStyleOriginalOverride?: ChartScatterStyle;
  scatterStyleOverride?: ChartScatterStyle;
  scatterStyleChanged: boolean;
  seriesDirectionBaseline: ChartSeriesDirection;
  seriesDirectionDirty: boolean;
  title: string;
  /** Excel's automatic title, which the text field does not override. */
  titleAuto: boolean;
  legendVisible: boolean;
  legendBaseline: boolean;
  legendOriginalOverride?: boolean;
  legendOverride?: boolean;
  legendManuallyChanged: boolean;
  legendKindChanged: boolean;
  legendPosition: 'top' | 'right' | 'bottom' | 'left';
  grouping: ChartGrouping;
  seriesDirection: ChartSeriesDirection;
  emptyCells: ChartEmptyCellMode;
  xAxis: ChartAxisDraft;
  yAxis: ChartAxisDraft;
  xAxisBaseline: ChartAxisDraft;
  yAxisBaseline: ChartAxisDraft;
  xAxisChanged: Set<ChartAxisField>;
  yAxisChanged: Set<ChartAxisField>;
  gapWidth: number;
  overlap: number;
  firstSliceAngle: number;
  holeSize: number;
  holeSizeBaseline: number;
  holeSizeChanged: boolean;
  dataLabels: MutableDataLabels;
  dataLabelsBaseline: MutableDataLabels;
  dataLabelsChanged: Set<keyof MutableDataLabels>;
  palette: string[];
  paletteBaseline: string[];
  varyColors: boolean;
  varyColorsBaseline: boolean;
  varyColorsOriginalOverride?: boolean;
  varyColorsOverride?: boolean;
  varyColorsManuallyChanged: boolean;
  varyColorsKindChanged: boolean;
  filter: { hiddenSeriesIds: string[]; hiddenCategoryIds: string[] };
  seriesStyles: ChartSeriesStyleSpec[];
}

type ChartAxisField =
  | 'visible'
  | 'reverseOrder'
  | 'crossing'
  | 'labelPosition'
  | 'logBase'
  | 'title'
  | 'titleAuto'
  | 'min'
  | 'max'
  | 'majorUnit'
  | 'numberFormat'
  | 'majorGridlines';

export type AxisCrossingMode = 'auto' | 'min' | 'max' | 'custom';
export type AxisLabelPositionMode = 'automatic' | 'nextTo' | 'low' | 'high' | 'none';

const validAxisLogBase = (value: unknown): number | undefined =>
  typeof value === 'number' && Number.isFinite(value) && value >= 2 && value <= 1000
    ? value
    : undefined;

const normalizedAxisLogBase = (value: number): number => {
  if (!Number.isFinite(value)) return 10;
  return Math.max(2, Math.min(1000, value));
};

interface ChartAxisDraft {
  visible: boolean;
  reverseOrder: boolean;
  /** The raw value is retained until the crossing control is edited. */
  crossing?: ChartAxisSpec['crossing'];
  crossingMode: AxisCrossingMode;
  crossingCustom: number;
  /** The raw value is retained until the label position control is edited. */
  labelPosition?: ChartAxisSpec['labelPosition'];
  labelPositionMode: AxisLabelPositionMode;
  /** The raw value is retained until the logarithmic-scale control is edited. */
  logBase?: ChartAxisSpec['logBase'];
  logEnabled: boolean;
  logBaseDraft: number;
  title: string;
  /** Excel's automatic title, which the text field does not override. */
  titleAuto: boolean;
  min?: number;
  max?: number;
  majorUnit?: number;
  numberFormat: string;
  majorGridlines: boolean;
}

const axisFromChart = (axis: SessionChartSettingsShape['xAxis']): ChartAxisDraft => {
  const raw = (axis ?? {}) as Record<string, unknown>;
  const finite = (value: unknown): number | undefined =>
    typeof value === 'number' && Number.isFinite(value) ? value : undefined;
  const crossing = raw.crossing as ChartAxisSpec['crossing'];
  const crossingMode: AxisCrossingMode =
    crossing === 'min' || crossing === 'max' || crossing === 'auto'
      ? crossing
      : finite(crossing) === undefined
        ? 'auto'
        : 'custom';
  const labelPosition = raw.labelPosition as ChartAxisSpec['labelPosition'];
  const labelPositionMode: AxisLabelPositionMode =
    labelPosition === 'nextTo' ||
    labelPosition === 'low' ||
    labelPosition === 'high' ||
    labelPosition === 'none'
      ? labelPosition
      : 'automatic';
  const logBase = raw.logBase as ChartAxisSpec['logBase'];
  const validLogBase = validAxisLogBase(logBase);
  return {
    visible: raw.visible !== false,
    reverseOrder: raw.reverseOrder === true,
    crossing,
    crossingMode,
    crossingCustom: finite(crossing) ?? 0,
    labelPosition,
    labelPositionMode,
    logBase,
    logEnabled: validLogBase !== undefined,
    logBaseDraft: validLogBase ?? 10,
    title: typeof raw.title === 'string' ? raw.title : '',
    titleAuto: typeof raw.title === 'object',
    min: finite(raw.min),
    max: finite(raw.max),
    majorUnit: (() => {
      const value = finite(raw.majorUnit);
      return value !== undefined && value > 0 ? value : undefined;
    })(),
    numberFormat: typeof raw.numberFormat === 'string' ? raw.numberFormat : '',
    majorGridlines: raw.majorGridlines !== false,
  };
};

export type MutableDataLabels = {
  -readonly [Key in keyof ChartDataLabelSpec]-?: ChartDataLabelSpec[Key];
};

const defaultDataLabels = (): MutableDataLabels => ({
  visible: false,
  showSeriesName: false,
  showCategoryName: false,
  showValue: true,
  showPercentage: false,
  showLegendKey: false,
  separator: ', ',
  position: 'best-fit',
  numberFormat: '',
  color: '#1f1f1f',
  fontSize: 10,
});

const cloneDataLabels = (value?: ChartDataLabelSpec): MutableDataLabels => {
  const next = { ...defaultDataLabels(), ...value };
  if (
    value?.showValue === undefined &&
    [
      value?.showSeriesName,
      value?.showCategoryName,
      value?.showPercentage,
      value?.showLegendKey,
    ].some((flag) => flag !== undefined)
  ) {
    next.showValue = false;
  }
  next.fontSize = Number.isFinite(value?.fontSize)
    ? Math.max(6, Math.min(72, value?.fontSize ?? 10))
    : 10;
  return next;
};

/** The colours the series are drawn with: the palette when one is set, else each
 *  series' own colour, which is how the workbook stores them. */
const effectivePalette = (
  chart: SessionChartSettingsShape,
  data: ChartData,
  automatic: readonly string[],
): string[] => {
  if (chart.palette && chart.palette.length > 0) return [...chart.palette];
  const fills = data.series.map((series) =>
    seriesColorOf(
      chart,
      chart.seriesStyles?.find((style) => style.seriesId === series.id),
    ),
  );
  return Array.from(
    { length: Math.max(6, fills.length) },
    (_, index) =>
      fills[index] ?? automatic[index % automatic.length] ?? DEFAULT_THEME_ACCENTS[0] ?? '',
  );
};

/** Drop the series' own colours so an edited palette takes effect. */
export const releaseSeriesColors = (draft: ChartDraft): void => {
  draft.seriesStyles = draft.seriesStyles.map((style) => withoutSeriesColor(draft.kind, style));
};

export const isCircularKind = (kind: string): boolean => kind === 'pie' || kind === 'doughnut';
export const isDoughnutKind = (kind: string): boolean => kind === 'doughnut';

export const defaultVaryColors = (kind: string): boolean => isCircularKind(kind);

const scatterStyleOf = (chart: SessionChartSettingsShape): ChartScatterStyle | undefined => {
  const value = chart.scatterStyle;
  return value === 'marker' || value === 'line' || value === 'lineMarker' ? value : undefined;
};

const effectiveLegendVisible = (chart: SessionChartSettingsShape, data: ChartData): boolean => {
  if (chart.legend?.visible !== undefined) return chart.legend.visible;
  const prepared = prepareChartData(data, chart);
  return prepared.data.domain === 'category' && isCircularKind(chart.kind)
    ? prepared.data.categories.length > 1 && prepared.data.series.length > 0
    : prepared.data.series.length > 1;
};

const boundedHoleSize = (value: number | undefined): number =>
  Number.isFinite(value) ? Math.max(10, Math.min(90, value ?? 50)) : 50;

/**
 * Build the transient chart spec used by settings controls. This lets every
 * panel answer questions about the draft chart family and data orientation
 * without writing defaults into the persisted chart.
 */
export const effectiveDraftChart = (
  chart: SessionChartSettingsShape,
  draft: ChartDraft,
): SessionChartSettingsShape =>
  ({
    ...chart,
    kind: draft.kind,
    scatterStyle: draft.scatterStyle,
    grouping: draft.grouping,
    seriesDirection: draft.seriesDirection,
    emptyCells: draft.emptyCells,
    holeSize: draft.holeSize,
    palette: [...draft.palette],
    varyColors: draft.varyColors,
    legend:
      draft.legendOverride === undefined
        ? chart.legend
          ? { ...chart.legend, visible: undefined }
          : undefined
        : { ...(chart.legend ?? {}), visible: draft.legendOverride },
    filter: {
      hiddenSeriesIds: [...draft.filter.hiddenSeriesIds],
      hiddenCategoryIds: [...draft.filter.hiddenCategoryIds],
    },
    xAxis: { ...draft.xAxis },
    yAxis: { ...draft.yAxis },
  }) as SessionChartSettingsShape;

export const effectiveLegendDefault = (
  chart: SessionChartSettingsShape,
  draft: ChartDraft,
  data: ChartData,
): boolean => {
  const draftSpec = effectiveDraftChart(chart, draft);
  const spec = {
    ...draftSpec,
    legend: draftSpec.legend ? { ...draftSpec.legend, visible: undefined } : undefined,
  } as SessionChartSettingsShape;
  return effectiveLegendVisible(spec, data);
};

export const draftFromChart = (
  chart: SessionChartSettingsShape,
  data: ChartData,
  automatic: readonly string[],
): ChartDraft => {
  const palette = effectivePalette(chart, data, automatic);
  const dataLabels = cloneDataLabels(chart.dataLabels);
  const legendBaseline = effectiveLegendVisible(chart, data);
  const legendOriginalOverride = chart.legend?.visible;
  const varyColorsBaseline = chart.varyColors ?? defaultVaryColors(chart.kind);
  const holeSizeBaseline = boundedHoleSize(chart.holeSize);
  const scatterStyleOriginalOverride = scatterStyleOf(chart);
  const scatterStyleBaseline = scatterStyleOriginalOverride ?? 'marker';
  const xAxis = axisFromChart(chart.xAxis);
  const yAxis = axisFromChart(chart.yAxis);
  const categoryCount =
    data.domain === 'category'
      ? data.categories.reduce((count, _category, index) => {
          const id = data.categoryIds?.[index] ?? `category-${index}`;
          return chart.filter?.hiddenCategoryIds?.includes(id) ? count : count + 1;
        }, 0)
      : 0;
  xAxis.crossingCustom = normalizedCrossingCustom(
    xAxis.crossingCustom,
    axisIsNumeric(chart, 'xAxis'),
    categoryCount,
  );
  yAxis.crossingCustom = normalizedCrossingCustom(
    yAxis.crossingCustom,
    axisIsNumeric(chart, 'yAxis'),
    categoryCount,
  );
  return {
    kind: chart.kind,
    kindBaseline: chart.kind,
    kindDirty: false,
    kindChanged: false,
    scatterStyle: scatterStyleBaseline,
    scatterStyleBaseline,
    ...(scatterStyleOriginalOverride === undefined
      ? {}
      : {
          scatterStyleOverride: scatterStyleOriginalOverride,
          scatterStyleOriginalOverride,
        }),
    scatterStyleChanged: false,
    seriesDirectionBaseline: chart.seriesDirection ?? 'auto',
    seriesDirectionDirty: false,
    title: typeof chart.title === 'string' ? chart.title : '',
    titleAuto: typeof chart.title === 'object',
    legendVisible: legendBaseline,
    legendBaseline,
    ...(legendOriginalOverride === undefined ? {} : { legendOverride: legendOriginalOverride }),
    ...(legendOriginalOverride === undefined ? {} : { legendOriginalOverride }),
    legendManuallyChanged: false,
    legendKindChanged: false,
    legendPosition: chart.legend?.position ?? 'right',
    grouping: chart.grouping ?? 'clustered',
    seriesDirection: chart.seriesDirection ?? 'auto',
    emptyCells: chart.emptyCells ?? 'gap',
    xAxis,
    yAxis,
    xAxisBaseline: { ...xAxis },
    yAxisBaseline: { ...yAxis },
    xAxisChanged: new Set(),
    yAxisChanged: new Set(),
    gapWidth: boundedGapWidth(chart.gapWidth),
    overlap: boundedOverlap(chart.overlap),
    firstSliceAngle: boundedFirstSliceAngle(chart.firstSliceAngle),
    holeSize: holeSizeBaseline,
    holeSizeBaseline,
    holeSizeChanged: false,
    dataLabels,
    dataLabelsBaseline: { ...dataLabels },
    dataLabelsChanged: new Set(),
    palette: [...palette],
    paletteBaseline: palette,
    varyColors: varyColorsBaseline,
    varyColorsBaseline,
    ...(chart.varyColors === undefined ? {} : { varyColorsOverride: chart.varyColors }),
    ...(chart.varyColors === undefined ? {} : { varyColorsOriginalOverride: chart.varyColors }),
    varyColorsManuallyChanged: false,
    varyColorsKindChanged: false,
    filter: {
      hiddenSeriesIds: [...(chart.filter?.hiddenSeriesIds ?? [])],
      hiddenCategoryIds: [...(chart.filter?.hiddenCategoryIds ?? [])],
    },
    seriesStyles: cloneStyles(chart.seriesStyles) ?? [],
  };
};

export const updateDataLabel = <Key extends keyof MutableDataLabels>(
  draft: ChartDraft,
  key: Key,
  value: MutableDataLabels[Key],
): void => {
  draft.dataLabels[key] = value;
  if (jsonEqual(value, draft.dataLabelsBaseline[key])) draft.dataLabelsChanged.delete(key);
  else draft.dataLabelsChanged.add(key);
};

const jsonEqual = (left: unknown, right: unknown): boolean =>
  JSON.stringify(left) === JSON.stringify(right);

const boundedGapWidth = (value: number | undefined): number =>
  Number.isFinite(value) ? Math.max(0, Math.min(500, value ?? 150)) : 150;

const boundedOverlap = (value: number | undefined): number =>
  Number.isFinite(value) ? Math.max(-100, Math.min(100, value ?? 0)) : 0;

const boundedFirstSliceAngle = (value: number | undefined): number =>
  Number.isFinite(value) ? Math.max(0, Math.min(360, value ?? 0)) : 0;

export const normalizedCrossingCustom = (
  value: number,
  numeric: boolean,
  categoryCount: number,
): number => {
  if (numeric) return Number.isFinite(value) ? value : 0;
  const count = Math.max(1, categoryCount);
  const rounded = Number.isFinite(value) ? Math.round(value) : 1;
  return Math.max(1, Math.min(count, rounded));
};

export const categoryCountForDraft = (data: ChartData, draft: ChartDraft): number => {
  if (data.domain !== 'category') return 0;
  return data.categories.reduce((count, _category, index) => {
    const id = data.categoryIds?.[index] ?? `category-${index}`;
    return draft.filter.hiddenCategoryIds.includes(id) ? count : count + 1;
  }, 0);
};

export const updateAxisField = <Key extends ChartAxisField>(
  draft: ChartDraft,
  axisName: 'xAxis' | 'yAxis',
  key: Key,
  value: ChartAxisDraft[Key],
): void => {
  const axis = draft[axisName];
  const baseline = draft[axisName === 'xAxis' ? 'xAxisBaseline' : 'yAxisBaseline'];
  const changed = draft[axisName === 'xAxis' ? 'xAxisChanged' : 'yAxisChanged'];
  axis[key] = value as never;
  if (jsonEqual(value, baseline[key])) changed.delete(key);
  else changed.add(key);
};

const crossingMatchesBaseline = (axis: ChartAxisDraft, baseline: ChartAxisDraft): boolean =>
  axis.crossingMode === baseline.crossingMode &&
  (axis.crossingMode !== 'custom' || axis.crossingCustom === baseline.crossingCustom);

const updateCrossingChanged = (draft: ChartDraft, axisName: AxisName): void => {
  const axis = draft[axisName];
  const baseline = draft[axisName === 'xAxis' ? 'xAxisBaseline' : 'yAxisBaseline'];
  const changed = draft[axisName === 'xAxis' ? 'xAxisChanged' : 'yAxisChanged'];
  if (crossingMatchesBaseline(axis, baseline)) {
    changed.delete('crossing');
  } else {
    changed.add('crossing');
  }
};

export const updateAxisCrossingMode = (
  draft: ChartDraft,
  axisName: AxisName,
  mode: AxisCrossingMode,
  numeric: boolean,
  categoryCount: number,
): void => {
  const axis = draft[axisName];
  axis.crossingMode = mode;
  if (mode === 'auto') axis.crossing = undefined;
  else if (mode === 'min' || mode === 'max') axis.crossing = mode;
  else {
    axis.crossingCustom = normalizedCrossingCustom(axis.crossingCustom, numeric, categoryCount);
    axis.crossing = axis.crossingCustom;
  }
  updateCrossingChanged(draft, axisName);
};

export const updateAxisCrossingCustom = (
  draft: ChartDraft,
  axisName: AxisName,
  value: string,
  numeric: boolean,
  categoryCount: number,
): void => {
  const parsed = Number(value);
  const fallback = numeric ? 0 : 1;
  const next = normalizedCrossingCustom(
    Number.isFinite(parsed) ? parsed : fallback,
    numeric,
    categoryCount,
  );
  const axis = draft[axisName];
  axis.crossingCustom = next;
  if (axis.crossingMode === 'custom') axis.crossing = next;
  updateCrossingChanged(draft, axisName);
};

const labelPositionMatchesBaseline = (axis: ChartAxisDraft, baseline: ChartAxisDraft): boolean =>
  axis.labelPositionMode === baseline.labelPositionMode;

const updateLabelPositionChanged = (draft: ChartDraft, axisName: AxisName): void => {
  const axis = draft[axisName];
  const baseline = draft[axisName === 'xAxis' ? 'xAxisBaseline' : 'yAxisBaseline'];
  const changed = draft[axisName === 'xAxis' ? 'xAxisChanged' : 'yAxisChanged'];
  if (labelPositionMatchesBaseline(axis, baseline)) changed.delete('labelPosition');
  else changed.add('labelPosition');
};

export const updateAxisLabelPosition = (
  draft: ChartDraft,
  axisName: AxisName,
  mode: AxisLabelPositionMode,
): void => {
  const axis = draft[axisName];
  axis.labelPositionMode = mode;
  axis.labelPosition = mode === 'automatic' ? undefined : mode;
  updateLabelPositionChanged(draft, axisName);
};

const logMatchesBaseline = (axis: ChartAxisDraft, baseline: ChartAxisDraft): boolean =>
  axis.logEnabled === baseline.logEnabled &&
  (!axis.logEnabled || axis.logBaseDraft === baseline.logBaseDraft);

const updateLogChanged = (draft: ChartDraft, axisName: AxisName): void => {
  const axis = draft[axisName];
  const baseline = draft[axisName === 'xAxis' ? 'xAxisBaseline' : 'yAxisBaseline'];
  const changed = draft[axisName === 'xAxis' ? 'xAxisChanged' : 'yAxisChanged'];
  if (logMatchesBaseline(axis, baseline)) changed.delete('logBase');
  else changed.add('logBase');
};

export const updateAxisLogEnabled = (
  draft: ChartDraft,
  axisName: AxisName,
  enabled: boolean,
): void => {
  const axis = draft[axisName];
  axis.logEnabled = enabled;
  if (enabled) {
    axis.logBaseDraft = normalizedAxisLogBase(axis.logBaseDraft);
    axis.logBase = axis.logBaseDraft;
  } else {
    axis.logBase = undefined;
  }
  updateLogChanged(draft, axisName);
};

export const updateAxisLogBase = (draft: ChartDraft, axisName: AxisName, value: string): void => {
  const trimmed = value.trim();
  const parsed = trimmed === '' ? Number.NaN : Number(trimmed);
  const axis = draft[axisName];
  axis.logBaseDraft = normalizedAxisLogBase(Number.isFinite(parsed) ? parsed : 10);
  if (axis.logEnabled) axis.logBase = axis.logBaseDraft;
  updateLogChanged(draft, axisName);
};

const axisPatch = (
  original: SessionChartSettingsShape['xAxis'],
  draft: ChartAxisDraft,
  changed: Set<ChartAxisField>,
): Record<string, unknown> | undefined => {
  if (changed.size === 0) return undefined;
  const next = { ...((original ?? {}) as Record<string, unknown>) };
  for (const key of changed) {
    if (key === 'crossing') {
      if (draft.crossingMode === 'auto') delete next.crossing;
      else if (draft.crossingMode === 'min' || draft.crossingMode === 'max') {
        next.crossing = draft.crossingMode;
      } else {
        next.crossing = draft.crossingCustom;
      }
    } else if (key === 'labelPosition') {
      if (draft.labelPositionMode === 'automatic') delete next.labelPosition;
      else next.labelPosition = draft.labelPositionMode;
    } else if (key === 'logBase') {
      if (!draft.logEnabled) delete next.logBase;
      else next.logBase = draft.logBaseDraft;
    } else if (key === 'title' || key === 'titleAuto') {
      next.title = draft.titleAuto ? { auto: true } : draft.title.trim() || undefined;
    } else if (key === 'numberFormat') next[key] = draft.numberFormat.trim() || undefined;
    else next[key] = draft[key];
  }
  return next;
};

export const toPatch = (
  draft: ChartDraft,
  chart: SessionChartSettingsShape,
): SessionChartSettingsPatch => {
  const patch: SessionChartSettingsPatch = {};
  if (draft.kind !== chart.kind) patch.kind = draft.kind;
  if (draft.scatterStyleChanged && draft.scatterStyleOverride !== scatterStyleOf(chart)) {
    patch.scatterStyle = draft.scatterStyleOverride;
  }
  if (draft.titleAuto) {
    if (typeof chart.title !== 'object') patch.title = { auto: true };
  } else {
    const title = draft.title.trim() || undefined;
    if (title !== (typeof chart.title === 'string' ? chart.title : undefined)) patch.title = title;
  }

  const currentLegendPosition = chart.legend?.position ?? 'right';
  const currentLegendOverride = chart.legend?.visible;
  const legendChanged = draft.legendOverride !== currentLegendOverride;
  if (legendChanged || draft.legendPosition !== currentLegendPosition) {
    const legend: Record<string, unknown> = { ...(chart.legend ?? {}) };
    if (legendChanged) {
      if (draft.legendOverride === undefined) delete legend.visible;
      else legend.visible = draft.legendOverride;
    }
    if (draft.legendPosition !== currentLegendPosition) legend.position = draft.legendPosition;
    patch.legend = legend as SessionChartSettingsPatch['legend'];
  }

  if (draft.grouping !== (chart.grouping ?? 'clustered')) patch.grouping = draft.grouping;
  if (draft.seriesDirection !== (chart.seriesDirection ?? 'auto')) {
    patch.seriesDirection = draft.seriesDirection;
  }
  if (draft.emptyCells !== (chart.emptyCells ?? 'gap')) patch.emptyCells = draft.emptyCells;

  const xAxis = axisPatch(chart.xAxis, draft.xAxis, draft.xAxisChanged);
  if (xAxis) patch.xAxis = xAxis;
  const yAxis = axisPatch(chart.yAxis, draft.yAxis, draft.yAxisChanged);
  if (yAxis) patch.yAxis = yAxis;

  if (draft.gapWidth !== boundedGapWidth(chart.gapWidth)) patch.gapWidth = draft.gapWidth;
  if (draft.overlap !== boundedOverlap(chart.overlap)) patch.overlap = draft.overlap;
  if (draft.firstSliceAngle !== boundedFirstSliceAngle(chart.firstSliceAngle)) {
    patch.firstSliceAngle = draft.firstSliceAngle;
  }
  if (draft.holeSizeChanged && draft.holeSize !== chart.holeSize) {
    patch.holeSize = draft.holeSize;
  }

  if (draft.dataLabelsChanged.size > 0) {
    const dataLabels: Record<string, unknown> = { ...(chart.dataLabels ?? {}) };
    for (const key of draft.dataLabelsChanged) dataLabels[key] = draft.dataLabels[key];
    const valueDefaultsToImplicit =
      chart.dataLabels?.showValue === undefined &&
      draft.dataLabels.showValue &&
      (['showSeriesName', 'showCategoryName', 'showPercentage', 'showLegendKey'] as const).some(
        (key) => draft.dataLabelsChanged.has(key),
      );
    if (valueDefaultsToImplicit) dataLabels.showValue = true;
    patch.dataLabels = dataLabels as ChartDataLabelSpec;
  }
  if (!jsonEqual(draft.palette, draft.paletteBaseline)) patch.palette = [...draft.palette];
  const currentVaryColorsOverride = chart.varyColors;
  const varyColorsChanged = draft.varyColorsOverride !== currentVaryColorsOverride;
  if (varyColorsChanged) {
    patch.varyColors = draft.varyColorsOverride;
  }

  const filter = {
    hiddenSeriesIds: [...draft.filter.hiddenSeriesIds],
    hiddenCategoryIds: [...draft.filter.hiddenCategoryIds],
  };
  const currentFilter = {
    hiddenSeriesIds: [...(chart.filter?.hiddenSeriesIds ?? [])],
    hiddenCategoryIds: [...(chart.filter?.hiddenCategoryIds ?? [])],
  };
  if (!jsonEqual(filter, currentFilter)) patch.filter = filter;

  const seriesStyles = cloneStyles(draft.seriesStyles) ?? [];
  if (!jsonEqual(seriesStyles, chart.seriesStyles ?? [])) patch.seriesStyles = seriesStyles;
  return patch;
};

export const seriesEntries = (
  data: ChartData,
  labels?: SessionChartSettingsLabels,
): readonly { id: string; label: string }[] =>
  data.series.map((series, index) => ({
    id: series.id,
    label: series.name ?? `${labels?.series ?? DEFAULT_LABELS.series} ${index + 1}`,
  }));

export const styleForSeries = (draft: ChartDraft, seriesId: string): ChartSeriesStyleSpec =>
  draft.seriesStyles.find((style) => style.seriesId === seriesId) ?? { seriesId };

export const updateSeriesStyle = (
  draft: ChartDraft,
  seriesId: string,
  update: Partial<ChartSeriesStyleSpec>,
): void => {
  const current = styleForSeries(draft, seriesId);
  const next = { ...current, ...update, seriesId };
  const index = draft.seriesStyles.findIndex((style) => style.seriesId === seriesId);
  if (index < 0) draft.seriesStyles.push(next);
  else draft.seriesStyles[index] = next;
};

const updatePointStyle = (
  draft: ChartDraft,
  seriesId: string,
  pointId: string,
  update: Partial<ChartPointStyleSpec>,
): void => {
  const current = styleForSeries(draft, seriesId);
  const points = [...(current.points ?? [])];
  const index = points.findIndex((point) => point.pointId === pointId);
  const next = {
    ...(index < 0 ? { pointId } : points[index]),
    ...update,
    pointId,
  };
  if (index < 0) points.push(next);
  else points[index] = next;
  updateSeriesStyle(draft, seriesId, { points });
};

export const patchFill = (
  fill: ChartFillSpec | undefined,
  patch: Partial<ChartFillSpec>,
): ChartFillSpec => ({
  ...(fill ?? {}),
  ...patch,
});

export const patchLine = (
  line: ChartLineSpec | undefined,
  patch: Partial<ChartLineSpec>,
): ChartLineSpec => ({
  ...(line ?? {}),
  ...patch,
});

export type AxisName = 'xAxis' | 'yAxis';

export const axisIsNumeric = (
  chart: Pick<SessionChartSettingsShape, 'kind'>,
  axis: AxisName,
): boolean => {
  if (isCircularKind(chart.kind)) return false;
  if (chart.kind === 'scatter') return true;
  if (chart.kind === 'bar') return axis === 'xAxis';
  return axis === 'yAxis';
};

export const stylePointEntries = (
  data: ChartData,
  seriesId: string,
  labels?: SessionChartSettingsLabels,
): readonly { id: string; label: string }[] => {
  if (data.domain === 'category') {
    return data.categories.map((label, index) => ({
      id: data.categoryIds?.[index] ?? `category-${index}`,
      label: label || `${labels?.categories ?? DEFAULT_LABELS.categories} ${index + 1}`,
    }));
  }
  const series = data.series.find((candidate) => candidate.id === seriesId);
  return (series?.points ?? []).map((point, index) => ({
    id: point.id ?? `point-${index}`,
    label: `${labels?.point ?? DEFAULT_LABELS.point} ${index + 1}`,
  }));
};

export const styleTarget = (
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
  update: Partial<ChartSeriesStyleSpec>,
): void => {
  if (pointId === null) updateSeriesStyle(draft, seriesId, update);
  else updatePointStyle(draft, seriesId, pointId, update);
};

export const syncUntouchedOptionalOverrides = (
  draft: ChartDraft,
  chart: SessionChartSettingsShape,
  data: ChartData,
): void => {
  for (const axisName of ['xAxis', 'yAxis'] as const) {
    const changed = draft[axisName === 'xAxis' ? 'xAxisChanged' : 'yAxisChanged'];
    if (!changed.has('reverseOrder')) {
      const reverseOrder = axisFromChart(chart[axisName]).reverseOrder;
      draft[axisName].reverseOrder = reverseOrder;
      draft[axisName === 'xAxis' ? 'xAxisBaseline' : 'yAxisBaseline'].reverseOrder = reverseOrder;
    }

    const baseline = draft[axisName === 'xAxis' ? 'xAxisBaseline' : 'yAxisBaseline'];
    if (!changed.has('crossing')) {
      const nextCrossing = axisFromChart(chart[axisName]);
      nextCrossing.crossingCustom = normalizedCrossingCustom(
        nextCrossing.crossingCustom,
        axisIsNumeric({ kind: draft.kind }, axisName),
        categoryCountForDraft(data, draft),
      );
      draft[axisName].crossing = nextCrossing.crossing;
      draft[axisName].crossingMode = nextCrossing.crossingMode;
      draft[axisName].crossingCustom = nextCrossing.crossingCustom;
      baseline.crossing = nextCrossing.crossing;
      baseline.crossingMode = nextCrossing.crossingMode;
      baseline.crossingCustom = nextCrossing.crossingCustom;
    }

    if (!changed.has('labelPosition')) {
      const nextLabelPosition = axisFromChart(chart[axisName]);
      draft[axisName].labelPosition = nextLabelPosition.labelPosition;
      draft[axisName].labelPositionMode = nextLabelPosition.labelPositionMode;
      baseline.labelPosition = nextLabelPosition.labelPosition;
      baseline.labelPositionMode = nextLabelPosition.labelPositionMode;
    }

    if (!changed.has('logBase')) {
      const nextLog = axisFromChart(chart[axisName]);
      draft[axisName].logBase = nextLog.logBase;
      draft[axisName].logEnabled = nextLog.logEnabled;
      draft[axisName].logBaseDraft = nextLog.logBaseDraft;
      baseline.logBase = nextLog.logBase;
      baseline.logEnabled = nextLog.logEnabled;
      baseline.logBaseDraft = nextLog.logBaseDraft;
    }
  }
  if (!draft.legendManuallyChanged) {
    const nextLegendOverride = chart.legend?.visible;
    if (nextLegendOverride === undefined) {
      delete draft.legendOriginalOverride;
      delete draft.legendOverride;
    } else {
      draft.legendOriginalOverride = nextLegendOverride;
      draft.legendOverride = nextLegendOverride;
    }
    const effective = effectiveDraftChart(chart, draft);
    const implicit = {
      ...effective,
      legend: effective.legend ? { ...effective.legend, visible: undefined } : undefined,
    } as SessionChartSettingsShape;
    draft.legendBaseline = effectiveLegendVisible(implicit, data);
    draft.legendVisible = draft.legendOverride ?? draft.legendBaseline;
  }
  if (!draft.varyColorsManuallyChanged) {
    const nextVaryColorsOverride = chart.varyColors;
    if (nextVaryColorsOverride === undefined) {
      delete draft.varyColorsOriginalOverride;
      delete draft.varyColorsOverride;
    } else {
      draft.varyColorsOriginalOverride = nextVaryColorsOverride;
      draft.varyColorsOverride = nextVaryColorsOverride;
    }
    draft.varyColorsBaseline = nextVaryColorsOverride ?? defaultVaryColors(String(draft.kind));
    draft.varyColors = draft.varyColorsBaseline;
  }
  if (!draft.scatterStyleChanged) {
    const nextScatterStyleOverride = scatterStyleOf(chart);
    if (nextScatterStyleOverride === undefined) {
      delete draft.scatterStyleOriginalOverride;
      delete draft.scatterStyleOverride;
    } else {
      draft.scatterStyleOriginalOverride = nextScatterStyleOverride;
      draft.scatterStyleOverride = nextScatterStyleOverride;
    }
    draft.scatterStyleBaseline = nextScatterStyleOverride ?? 'marker';
    draft.scatterStyle = draft.scatterStyleBaseline;
  }
};

export const refreshDraftDefaults = (
  draft: ChartDraft,
  chart: SessionChartSettingsShape,
  data: ChartData,
): void => {
  const effective = effectiveDraftChart(chart, draft);
  if (draft.legendOverride === undefined) {
    draft.legendVisible = effectiveLegendVisible(effective, data);
  }
  if (draft.varyColorsOverride === undefined) {
    draft.varyColors = defaultVaryColors(String(draft.kind));
  }
};

/** Align the draft with series added or removed in the workbook. */
export const reconcileSeriesDraft = (draft: ChartDraft, chart: SessionChartSettingsShape): void => {
  const known = new Set(chart.data.series.map((series) => series.id));
  const styles = draft.seriesStyles.filter((style) => known.has(style.seriesId));
  for (const style of cloneStyles(chart.seriesStyles) ?? []) {
    if (!styles.some((entry) => entry.seriesId === style.seriesId)) styles.push(style);
  }
  draft.seriesStyles = styles;
  draft.filter.hiddenSeriesIds = draft.filter.hiddenSeriesIds.filter((id) => known.has(id));
};
