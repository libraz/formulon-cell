import { seriesColorOf, withoutSeriesColor } from '../charts/engine-adapter.js';
import { prepareChartData } from '../charts/filters.js';
import { chartLabelPositions } from '../charts/label-options.js';
import { automaticPalette, DEFAULT_THEME_ACCENTS } from '../charts/series-colors.js';
import { resolveChartStyle } from '../charts/styles.js';
import type {
  ChartAxisSpec,
  ChartData,
  ChartDataLabelSpec,
  ChartEmptyCellMode,
  ChartFillSpec,
  ChartGrouping,
  ChartKind,
  ChartLabelPosition,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartPointStyleSpec,
  ChartScatterStyle,
  ChartSeriesDirection,
  ChartSeriesStyleSpec,
  ChartSpec,
  ChartValueFormatter as CoreChartValueFormatter,
} from '../charts/types.js';
import type { History } from '../commands/history.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';

/** The host may use this hook to format values according to its workbook
 * profile. The renderer keeps the hook optional so the chart package remains
 * usable without a workbook engine. */
export type ChartValueFormatter = CoreChartValueFormatter;

/** A chart the settings surface can edit: one the renderer can draw. */
export type SessionChartSettingsShape = SessionChart & { kind: ChartKind };
export type SessionChartSettingsPatch = {
  -readonly [Key in keyof ChartSpec]?: ChartSpec[Key];
};

export interface SessionChartSettingsLabels {
  columnChart: string;
  barChart: string;
  lineChart: string;
  areaChart: string;
  pieChart: string;
  doughnutChart: string;
  scatterChart: string;
  chartType: string;
  scatterStyle: string;
  scatterStyleMarker: string;
  scatterStyleLine: string;
  scatterStyleLineMarker: string;
  holeSize: string;
  settings: string;
  elementsTab: string;
  stylesTab: string;
  filtersTab: string;
  closeSettings: string;
  apply: string;
  cancel: string;
  dataLabels: string;
  showDataLabels: string;
  showSeriesName: string;
  showCategoryName: string;
  showValue: string;
  showPercentage: string;
  showLegendKey: string;
  labelSeparator: string;
  labelPosition: string;
  positionBestFit: string;
  positionAutomatic: string;
  positionCenter: string;
  positionInsideBase: string;
  positionInsideEnd: string;
  positionOutsideEnd: string;
  positionAbove: string;
  positionBelow: string;
  positionLeft: string;
  positionRight: string;
  numberFormat: string;
  textColor: string;
  fontSize: string;
  series: string;
  point: string;
  allPoints: string;
  fillColor: string;
  lineColor: string;
  lineWidth: string;
  transparency: string;
  lineDash: string;
  dashSolid: string;
  dashDash: string;
  dashDot: string;
  dashDashDot: string;
  marker: string;
  markerVisible: string;
  markerShape: string;
  shapeCircle: string;
  shapeSquare: string;
  shapeDiamond: string;
  shapeTriangle: string;
  shapeX: string;
  shapeStar: string;
  shapePlus: string;
  shapeDot: string;
  shapeDash: string;
  smoothLine: string;
  titleAutomatic: string;
  markerSize: string;
  categories: string;
  selectAll: string;
  clearAll: string;
  noSeries: string;
  noCategories: string;
  pivotChartLocked: string;
  addSeries: string;
  removeSeries: string;
  addSeriesConfirm: string;
  seriesNameRef: string;
  seriesValuesRef: string;
  seriesXValuesRef: string;
  seriesYValuesRef: string;
  axisLabelsRef: string;
  invalidReference: string;
  lastSeriesLocked: string;
  palette: string;
  paletteOffice: string;
  paletteBlue: string;
  paletteMonochrome: string;
  varyColors: string;
  title: string;
  legend: string;
  legendPosition: string;
  grouping: string;
  groupingClustered: string;
  groupingStacked: string;
  groupingPercentStacked: string;
  seriesDirection: string;
  seriesDirectionAuto: string;
  seriesDirectionColumns: string;
  seriesDirectionRows: string;
  emptyCells: string;
  emptyCellsGap: string;
  emptyCellsZero: string;
  emptyCellsSpan: string;
  axisTitles: string;
  xAxis: string;
  yAxis: string;
  xAxisTitle: string;
  yAxisTitle: string;
  showAxis: string;
  reverseOrder: string;
  verticalAxisCrosses: string;
  horizontalAxisCrosses: string;
  labelPositionNextTo: string;
  labelPositionLow: string;
  labelPositionHigh: string;
  labelPositionNone: string;
  logarithmicScale: string;
  logBase: string;
  custom: string;
  customValue: string;
  minimum: string;
  maximum: string;
  majorUnit: string;
  axisNumberFormat: string;
  majorGridlines: string;
  automatic: string;
  gapWidth: string;
  overlap: string;
  firstSliceAngle: string;
  positionTop: string;
  positionBottom: string;
}

export interface SessionChartSettingsHandle {
  open(
    chart: SessionChartSettingsShape,
    data: ChartData,
    anchor: HTMLElement,
    tab?: SettingsTab,
  ): void;
  sync(chart: SessionChartSettingsShape, data: ChartData, anchor: HTMLElement): void;
  close(discard?: boolean): void;
  setLabels(next: Partial<SessionChartSettingsLabels>, locale?: string): void;
  detach(): void;
  isOpen(): boolean;
  chartId(): string | null;
}

export type SettingsTab = 'elements' | 'styles' | 'filters';

/** Cell references of a series to add, as typed; `values` are a scatter chart's Y values. */
export interface SessionChartSeriesRefs {
  name?: string;
  values: string;
  categories?: string;
  x?: string;
}

export interface SessionChartSettingsDeps {
  readonly host: HTMLElement;
  readonly store: SpreadsheetStore;
  readonly history?: History | null;
  readonly locale?: string;
  readonly formatValue?: ChartValueFormatter;
  readonly labels?: Partial<SessionChartSettingsLabels>;
  readonly canEdit?: (chart: SessionChartSettingsShape) => boolean;
  /** Rebuild source data when a draft changes chart family or orientation. */
  readonly readData?: (chart: SessionChartSettingsShape) => ChartData;
  readonly onApply: (id: string, patch: SessionChartSettingsPatch, history: History | null) => void;
  /** Append a series to the workbook chart; its id, or null when refused. */
  readonly onAddSeries: (
    id: string,
    refs: SessionChartSeriesRefs,
    history: History | null,
  ) => string | null;
  /** Remove a series from the workbook chart; false when refused. */
  readonly onRemoveSeries: (id: string, seriesId: string, history: History | null) => boolean;
  /** The workbook theme's accent1..accent6; the default theme's when omitted. */
  readonly accents?: () => readonly string[] | undefined;
}

const SVG_SAFE_HEX = /^#[0-9a-f]{6}$/i;
const ALL_POINTS = '__all__';
/** Fixed presets; `office` is the theme's accents, the default theme's here. */
const PALETTE_PRESETS: Record<string, readonly string[]> = {
  office: DEFAULT_THEME_ACCENTS,
  blue: ['#1f4e78', '#5b9bd5', '#9dc3e6', '#d9eaf7', '#17365d', '#2f75b5'],
  mono: ['#1f1f1f', '#4d4d4d', '#737373', '#999999', '#bfbfbf', '#d9d9d9'],
};

const PALETTE_LABEL_KEYS: Record<string, keyof SessionChartSettingsLabels> = {
  office: 'paletteOffice',
  blue: 'paletteBlue',
  mono: 'paletteMonochrome',
};

const DEFAULT_LABELS: SessionChartSettingsLabels = {
  columnChart: 'Column chart',
  barChart: 'Bar chart',
  lineChart: 'Line chart',
  areaChart: 'Area chart',
  pieChart: 'Pie chart',
  doughnutChart: 'Doughnut chart',
  scatterChart: 'Scatter chart',
  chartType: 'Chart type',
  scatterStyle: 'Scatter style',
  scatterStyleMarker: 'Markers',
  scatterStyleLine: 'Straight lines',
  scatterStyleLineMarker: 'Straight lines with markers',
  holeSize: 'Hole size (%)',
  settings: 'Chart settings',
  elementsTab: 'Elements',
  stylesTab: 'Styles',
  filtersTab: 'Filters',
  closeSettings: 'Close chart settings',
  apply: 'Apply',
  cancel: 'Cancel',
  dataLabels: 'Data labels',
  showDataLabels: 'Show data labels',
  showSeriesName: 'Series name',
  showCategoryName: 'Category name',
  showValue: 'Value',
  showPercentage: 'Percentage',
  showLegendKey: 'Legend key',
  labelSeparator: 'Separator',
  labelPosition: 'Label position',
  positionBestFit: 'Best fit',
  positionAutomatic: 'Automatic',
  positionCenter: 'Center',
  positionInsideBase: 'Inside base',
  positionInsideEnd: 'Inside end',
  positionOutsideEnd: 'Outside end',
  positionAbove: 'Above',
  positionBelow: 'Below',
  positionLeft: 'Left',
  positionRight: 'Right',
  numberFormat: 'Number format',
  textColor: 'Text color',
  fontSize: 'Font size',
  series: 'Series',
  point: 'Point',
  allPoints: 'All points',
  fillColor: 'Fill color',
  lineColor: 'Line color',
  lineWidth: 'Line width',
  transparency: 'Transparency',
  lineDash: 'Line dash',
  dashSolid: 'Solid',
  dashDash: 'Dash',
  dashDot: 'Dot',
  dashDashDot: 'Dash-dot',
  marker: 'Marker',
  markerVisible: 'Visible',
  markerShape: 'Shape',
  shapeCircle: 'Circle',
  shapeSquare: 'Square',
  shapeDiamond: 'Diamond',
  shapeTriangle: 'Triangle',
  shapeX: 'X',
  shapeStar: 'Star',
  shapePlus: 'Plus',
  shapeDot: 'Dot',
  shapeDash: 'Dash',
  smoothLine: 'Smoothed line',
  titleAutomatic: 'Automatic title',
  markerSize: 'Size',
  categories: 'Categories',
  selectAll: 'Select all',
  clearAll: 'Clear all',
  noSeries: 'No series',
  noCategories: 'No categories',
  pivotChartLocked: 'The series and data range of a PivotChart follow its PivotTable.',
  addSeries: 'Add series',
  removeSeries: 'Remove series',
  addSeriesConfirm: 'Add',
  seriesNameRef: 'Series name',
  seriesValuesRef: 'Series values',
  seriesXValuesRef: 'Series X values',
  seriesYValuesRef: 'Series Y values',
  axisLabelsRef: 'Axis label range',
  invalidReference: 'Enter a valid cell reference.',
  lastSeriesLocked: 'A chart needs at least one series.',
  palette: 'Palette',
  paletteOffice: 'Office',
  paletteBlue: 'Blue',
  paletteMonochrome: 'Monochrome',
  varyColors: 'Vary colors by point',
  title: 'Title',
  legend: 'Legend',
  legendPosition: 'Legend position',
  grouping: 'Grouping',
  groupingClustered: 'Clustered',
  groupingStacked: 'Stacked',
  groupingPercentStacked: '100% stacked',
  seriesDirection: 'Series direction',
  seriesDirectionAuto: 'Automatic',
  seriesDirectionColumns: 'Columns',
  seriesDirectionRows: 'Rows',
  emptyCells: 'Empty cells',
  emptyCellsGap: 'Gaps',
  emptyCellsZero: 'Zero',
  emptyCellsSpan: 'Connect data points',
  axisTitles: 'Axis titles',
  xAxis: 'X-axis',
  yAxis: 'Y-axis',
  xAxisTitle: 'X-axis title',
  yAxisTitle: 'Y-axis title',
  showAxis: 'Show axis',
  reverseOrder: 'Reverse order',
  verticalAxisCrosses: 'Vertical axis crosses',
  horizontalAxisCrosses: 'Horizontal axis crosses',
  labelPositionNextTo: 'Next to axis',
  labelPositionLow: 'Low (bottom/left)',
  labelPositionHigh: 'High (top/right)',
  labelPositionNone: 'None',
  logarithmicScale: 'Logarithmic scale',
  logBase: 'Log base',
  custom: 'Custom',
  customValue: 'Custom value',
  minimum: 'Minimum',
  maximum: 'Maximum',
  majorUnit: 'Major unit',
  axisNumberFormat: 'Number format',
  majorGridlines: 'Major gridlines',
  automatic: 'Automatic',
  gapWidth: 'Gap width (%)',
  overlap: 'Overlap (%)',
  firstSliceAngle: 'First slice angle (°)',
  positionTop: 'Top',
  positionBottom: 'Bottom',
};

const labelsWithDefaults = (
  labels?: Partial<SessionChartSettingsLabels>,
): SessionChartSettingsLabels => ({ ...DEFAULT_LABELS, ...labels });

const supportsTextSelection = (
  value: Element | null,
): value is HTMLInputElement | HTMLTextAreaElement =>
  value instanceof HTMLTextAreaElement ||
  (value instanceof HTMLInputElement &&
    ['password', 'search', 'tel', 'text', 'url'].includes(value.type));

interface SettingsViewState {
  focusedIndex: number;
  selectionStart: number | null;
  selectionEnd: number | null;
  selectionDirection: 'forward' | 'backward' | 'none' | null;
  surfaceScrollTop: number;
  bodyScrollTop: number;
}

const captureSettingsView = (surface: HTMLElement): SettingsViewState => {
  const active = document.activeElement;
  const controls = Array.from(
    surface.querySelectorAll<HTMLElement>('button, input, select, textarea'),
  );
  const selectionTarget = supportsTextSelection(active) ? active : null;
  const body = surface.querySelector<HTMLElement>('.fc-chart-settings__body');
  return {
    focusedIndex: active instanceof HTMLElement ? controls.indexOf(active) : -1,
    selectionStart: selectionTarget?.selectionStart ?? null,
    selectionEnd: selectionTarget?.selectionEnd ?? null,
    selectionDirection: selectionTarget?.selectionDirection ?? null,
    surfaceScrollTop: surface.scrollTop,
    bodyScrollTop: body?.scrollTop ?? 0,
  };
};

const restoreSettingsView = (surface: HTMLElement, state: SettingsViewState): void => {
  const controls = Array.from(
    surface.querySelectorAll<HTMLElement>('button, input, select, textarea'),
  );
  const focused = controls[state.focusedIndex];
  if (focused) {
    focused.focus({ preventScroll: true });
    if (
      supportsTextSelection(focused) &&
      state.selectionStart !== null &&
      state.selectionEnd !== null
    ) {
      focused.setSelectionRange(
        state.selectionStart,
        state.selectionEnd,
        state.selectionDirection ?? 'none',
      );
    }
  }
  surface.scrollTop = state.surfaceScrollTop;
  const body = surface.querySelector<HTMLElement>('.fc-chart-settings__body');
  if (body) body.scrollTop = state.bodyScrollTop;
};

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

interface ChartDraft {
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

type AxisCrossingMode = 'auto' | 'min' | 'max' | 'custom';
type AxisLabelPositionMode = 'automatic' | 'nextTo' | 'low' | 'high' | 'none';

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

type MutableDataLabels = {
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
const releaseSeriesColors = (draft: ChartDraft): void => {
  draft.seriesStyles = draft.seriesStyles.map((style) => withoutSeriesColor(draft.kind, style));
};

const isCircularKind = (kind: string): boolean => kind === 'pie' || kind === 'doughnut';
const isDoughnutKind = (kind: string): boolean => kind === 'doughnut';

const defaultVaryColors = (kind: string): boolean => isCircularKind(kind);

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
const effectiveDraftChart = (
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

const effectiveLegendDefault = (
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

const draftFromChart = (
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

const updateDataLabel = <Key extends keyof MutableDataLabels>(
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

const normalizedCrossingCustom = (
  value: number,
  numeric: boolean,
  categoryCount: number,
): number => {
  if (numeric) return Number.isFinite(value) ? value : 0;
  const count = Math.max(1, categoryCount);
  const rounded = Number.isFinite(value) ? Math.round(value) : 1;
  return Math.max(1, Math.min(count, rounded));
};

const categoryCountForDraft = (data: ChartData, draft: ChartDraft): number => {
  if (data.domain !== 'category') return 0;
  return data.categories.reduce((count, _category, index) => {
    const id = data.categoryIds?.[index] ?? `category-${index}`;
    return draft.filter.hiddenCategoryIds.includes(id) ? count : count + 1;
  }, 0);
};

const updateAxisField = <Key extends ChartAxisField>(
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

const updateAxisCrossingMode = (
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

const updateAxisCrossingCustom = (
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

const updateAxisLabelPosition = (
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

const updateAxisLogEnabled = (draft: ChartDraft, axisName: AxisName, enabled: boolean): void => {
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

const updateAxisLogBase = (draft: ChartDraft, axisName: AxisName, value: string): void => {
  const trimmed = value.trim();
  const parsed = trimmed === '' ? Number.NaN : Number(trimmed);
  const axis = draft[axisName];
  axis.logBaseDraft = normalizedAxisLogBase(Number.isFinite(parsed) ? parsed : 10);
  if (axis.logEnabled) axis.logBase = axis.logBaseDraft;
  updateLogChanged(draft, axisName);
};

const parseOptionalNumber = (value: string): number | undefined => {
  if (value.trim() === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
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

const toPatch = (
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

const element = <K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
): HTMLElementTagNameMap[K] => {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
};

const textNode = (
  tag: 'span' | 'div' | 'h3' | 'legend',
  text: string,
  className?: string,
): HTMLElement => {
  const node = element(tag, className);
  node.textContent = text;
  return node;
};

const makeField = (parent: HTMLElement, label: string, control: HTMLElement): HTMLElement => {
  const row = element('label', 'fc-chart-settings__field');
  row.append(textNode('span', label, 'fc-chart-settings__label'), control);
  parent.appendChild(row);
  return row;
};

const makeCheckbox = (
  parent: HTMLElement,
  label: string,
  checked: boolean,
  onChange: (checked: boolean) => void,
  disabled = false,
  className = '',
): HTMLInputElement => {
  const row = element('label', `fc-chart-settings__check ${className}`.trim());
  const input = element('input');
  input.type = 'checkbox';
  input.checked = checked;
  projectDisabledState(input, disabled, null);
  input.addEventListener('change', () => onChange(input.checked));
  row.append(input, textNode('span', label));
  parent.appendChild(row);
  return input;
};

interface TextInputOptions {
  readonly allowNegative?: boolean;
  readonly min?: number;
  readonly max?: number;
  readonly step?: string;
  readonly placeholder?: string;
  readonly disabled?: boolean;
}

const makeTextInput = (
  parent: HTMLElement,
  label: string,
  value: string,
  onInput: (value: string) => void,
  type: 'text' | 'number' = 'text',
  options: TextInputOptions = {},
): HTMLInputElement => {
  const input = element('input');
  input.type = type;
  input.value = value;
  if (options.placeholder !== undefined) input.placeholder = options.placeholder;
  if (type === 'number') {
    input.inputMode = 'decimal';
    if (options.min !== undefined) input.min = String(options.min);
    else if (!options.allowNegative) input.min = '0';
    if (options.max !== undefined) input.max = String(options.max);
    if (options.step !== undefined) input.step = options.step;
  }
  projectDisabledState(input, options.disabled === true, null);
  input.addEventListener('input', () => onInput(input.value));
  makeField(parent, label, input);
  return input;
};

const makeSelect = (
  parent: HTMLElement,
  label: string,
  value: string,
  options: readonly { value: string; label: string; disabled?: boolean }[],
  onChange: (value: string) => void,
  disabled = false,
): HTMLSelectElement => {
  const select = element('select');
  for (const option of options) {
    const item = element('option');
    item.value = option.value;
    item.textContent = option.label;
    projectDisabledState(item, option.disabled === true, null);
    select.appendChild(item);
  }
  projectDisabledState(select, disabled, null);
  select.value = value;
  select.addEventListener('change', () => onChange(select.value));
  makeField(parent, label, select);
  return select;
};

const clampPercent = (value: string): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? Math.max(0, Math.min(100, parsed)) : 0;
};

const validColor = (value: string): string | undefined =>
  SVG_SAFE_HEX.test(value.trim()) ? value.trim().toLowerCase() : undefined;

const makeColorField = (
  parent: HTMLElement,
  label: string,
  value: string,
  onChange: (value: string) => void,
): void => {
  const row = element('div', 'fc-chart-settings__field fc-chart-settings__color-field');
  row.appendChild(textNode('span', label, 'fc-chart-settings__label'));
  const controls = element('span', 'fc-chart-settings__color-controls');
  const text = element('input');
  text.type = 'text';
  text.setAttribute('aria-label', label);
  text.inputMode = 'text';
  text.value = value;
  text.maxLength = 7;
  const native = element('input');
  native.type = 'color';
  native.setAttribute('aria-label', label);
  native.value = validColor(value) ?? DEFAULT_THEME_ACCENTS[0] ?? '#000000';
  const apply = (next: string): void => {
    const color = validColor(next);
    if (!color) return;
    text.value = color;
    native.value = color;
    onChange(color);
  };
  text.addEventListener('input', () => apply(text.value));
  native.addEventListener('input', () => apply(native.value));
  controls.append(text, native);
  row.appendChild(controls);
  parent.appendChild(row);
};

const seriesEntries = (
  data: ChartData,
  labels?: SessionChartSettingsLabels,
): readonly { id: string; label: string }[] =>
  data.series.map((series, index) => ({
    id: series.id,
    label: series.name ?? `${labels?.series ?? DEFAULT_LABELS.series} ${index + 1}`,
  }));

const categoryEntries = (
  data: ChartData,
  labels?: SessionChartSettingsLabels,
): readonly { id: string; label: string }[] => {
  if (data.domain === 'category') {
    return data.categories.map((category, index) => ({
      id: data.categoryIds?.[index] ?? `category-${index}`,
      label: category || `${labels?.categories ?? DEFAULT_LABELS.categories} ${index + 1}`,
    }));
  }
  const count = data.series.reduce((max, series) => Math.max(max, series.points.length), 0);
  return Array.from({ length: count }, (_, index) => ({
    id: `point-${index}`,
    label: `${labels?.point ?? DEFAULT_LABELS.point} ${index + 1}`,
  }));
};

const dataSignature = (data: ChartData): string =>
  JSON.stringify({
    domain: data.domain,
    series:
      data.domain === 'xy'
        ? data.series.map((series) => ({
            id: series.id,
            name: series.name,
            points: series.points.map((point) => point.id ?? ''),
          }))
        : data.series.map((series) => ({ id: series.id, name: series.name })),
    categories:
      data.domain === 'category'
        ? data.categories.map((category, index) => ({
            id: data.categoryIds?.[index] ?? `category-${index}`,
            name: category,
          }))
        : undefined,
  });

const styleForSeries = (draft: ChartDraft, seriesId: string): ChartSeriesStyleSpec =>
  draft.seriesStyles.find((style) => style.seriesId === seriesId) ?? { seriesId };

const updateSeriesStyle = (
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

const patchFill = (
  fill: ChartFillSpec | undefined,
  patch: Partial<ChartFillSpec>,
): ChartFillSpec => ({
  ...(fill ?? {}),
  ...patch,
});

const patchLine = (
  line: ChartLineSpec | undefined,
  patch: Partial<ChartLineSpec>,
): ChartLineSpec => ({
  ...(line ?? {}),
  ...patch,
});

const POSITION_LABELS: Record<ChartLabelPosition, keyof SessionChartSettingsLabels> = {
  'best-fit': 'positionBestFit',
  center: 'positionCenter',
  'inside-base': 'positionInsideBase',
  'inside-end': 'positionInsideEnd',
  'outside-end': 'positionOutsideEnd',
  above: 'positionAbove',
  below: 'positionBelow',
  left: 'positionLeft',
  right: 'positionRight',
};

const allowedPositions = (
  chart: Pick<SessionChartSettingsShape, 'kind' | 'grouping'>,
  current: ChartLabelPosition,
): readonly {
  value: ChartLabelPosition;
  label: keyof SessionChartSettingsLabels;
  disabled?: boolean;
}[] => {
  const supported = chartLabelPositions(chart);
  const options: {
    value: ChartLabelPosition;
    label: keyof SessionChartSettingsLabels;
    disabled?: boolean;
  }[] = isCircularKind(chart.kind)
    ? supported.map((value) => ({ value, label: POSITION_LABELS[value] }))
    : [
        {
          value: 'best-fit',
          label: 'positionAutomatic',
          disabled: chart.kind === 'area',
        },
        ...supported.map((value) => ({ value, label: POSITION_LABELS[value] })),
      ];
  // A workbook may already contain a position that is invalid for its current
  // chart grouping. Keep that value in the draft until the user chooses a
  // supported replacement, while leaving it unavailable as a new choice.
  if (!options.some((option) => option.value === current)) {
    options.push({ value: current, label: POSITION_LABELS[current], disabled: true });
  }
  return options;
};

const supportsPercentage = (
  chart: Pick<SessionChartSettingsShape, 'kind'>,
  grouping: ChartGrouping,
): boolean => isCircularKind(chart.kind) || grouping === 'percent-stacked';

const groupingSupported = (chart: Pick<SessionChartSettingsShape, 'kind'>): boolean =>
  chart.kind === 'bar' || chart.kind === 'column' || chart.kind === 'line' || chart.kind === 'area';

type AxisName = 'xAxis' | 'yAxis';

const axisIsNumeric = (chart: Pick<SessionChartSettingsShape, 'kind'>, axis: AxisName): boolean => {
  if (isCircularKind(chart.kind)) return false;
  if (chart.kind === 'scatter') return true;
  if (chart.kind === 'bar') return axis === 'xAxis';
  return axis === 'yAxis';
};

const axisLogEligible = (
  chart: Pick<SessionChartSettingsShape, 'kind' | 'grouping'>,
  axis: AxisName,
): boolean => {
  if (!axisIsNumeric(chart, axis)) return false;
  if (chart.kind === 'scatter') return true;
  return (chart.grouping ?? 'clustered') === 'clustered';
};

const emptyCellOptionDisabled = (
  chart: Pick<SessionChartSettingsShape, 'kind' | 'scatterStyle'>,
  grouping: ChartGrouping,
  value: ChartEmptyCellMode,
): boolean => {
  if (value !== 'span') return false;
  if (chart.kind === 'scatter') {
    return chart.scatterStyle !== 'line' && chart.scatterStyle !== 'lineMarker';
  }
  return chart.kind !== 'line' || grouping !== 'clustered';
};

const emptyCellsForcedZero = (
  chart: Pick<SessionChartSettingsShape, 'kind'>,
  grouping: ChartGrouping,
): boolean => (chart.kind === 'line' || chart.kind === 'area') && grouping !== 'clustered';

const stylePointEntries = (
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

const styleTarget = (
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
  update: Partial<ChartSeriesStyleSpec>,
): void => {
  if (pointId === null) updateSeriesStyle(draft, seriesId, update);
  else updatePointStyle(draft, seriesId, pointId, update);
};

const pointIdForData = (data: ChartData, seriesIndex: number, pointIndex: number): string => {
  if (data.domain === 'category') {
    return data.categoryIds?.[pointIndex] ?? `category-${pointIndex}`;
  }
  return data.series[seriesIndex]?.points[pointIndex]?.id ?? `point-${pointIndex}`;
};

const resolvedStyleForSelection = (
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
): ReturnType<typeof resolveChartStyle> => {
  const seriesIndex = Math.max(
    0,
    data.series.findIndex((series) => series.id === seriesId),
  );
  const selectedSeries = data.series[seriesIndex];
  let pointIndex: number | undefined;
  if (pointId !== null && selectedSeries) {
    const index =
      data.domain === 'category'
        ? data.categories.findIndex(
            (_, index) => pointIdForData(data, seriesIndex, index) === pointId,
          )
        : (data.series[seriesIndex]?.points.findIndex(
            (_point, index) => pointIdForData(data, seriesIndex, index) === pointId,
          ) ?? -1);
    pointIndex = index < 0 ? undefined : index;
  }
  return resolveChartStyle(
    {
      ...chart,
      scatterStyle: draft.scatterStyle,
      palette: draft.palette,
      seriesStyles: draft.seriesStyles,
      varyColors: draft.varyColors,
    },
    draft.palette,
    seriesId,
    pointId ?? undefined,
    seriesIndex,
    pointIndex,
  );
};

const renderColorAndLineControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
): void => {
  const style = styleForSeries(draft, seriesId);
  const point =
    pointId === null ? undefined : style.points?.find((entry) => entry.pointId === pointId);
  const resolved = resolvedStyleForSelection(chart, data, draft, seriesId, pointId);
  let currentFill = { ...(pointId === null ? (style.fill ?? {}) : (point?.fill ?? {})) };
  let currentLine = { ...(pointId === null ? (style.line ?? {}) : (point?.line ?? {})) };
  // Excel keeps no fill on a line or scatter series or point; its colour is the line's and markers'.
  if (chart.kind !== 'line' && chart.kind !== 'scatter') {
    makeColorField(body, labels.fillColor, resolved.fill?.color ?? resolved.color, (color) => {
      currentFill = patchFill(currentFill, { color });
      styleTarget(draft, seriesId, pointId, { fill: currentFill });
    });
    makeTextInput(
      body,
      `${labels.fillColor} ${labels.transparency}`,
      String(resolved.fill?.transparency ?? 0),
      (value) => {
        // A workbook fill carries its colour with its transparency, so the inherited colour is kept.
        currentFill = patchFill(currentFill, {
          color: currentFill.color ?? resolved.fill?.color ?? resolved.color,
          transparency: clampPercent(value),
        });
        styleTarget(draft, seriesId, pointId, { fill: currentFill });
      },
      'number',
    );
  }
  makeColorField(body, labels.lineColor, resolved.line?.color ?? resolved.color, (color) => {
    currentLine = patchLine(currentLine, { color });
    styleTarget(draft, seriesId, pointId, { line: currentLine });
  });
  makeTextInput(
    body,
    `${labels.lineColor} ${labels.transparency}`,
    String(resolved.line?.transparency ?? 0),
    (value) => {
      currentLine = patchLine(currentLine, {
        color: currentLine.color ?? resolved.line?.color ?? resolved.color,
        transparency: clampPercent(value),
      });
      styleTarget(draft, seriesId, pointId, { line: currentLine });
    },
    'number',
  );
  makeTextInput(
    body,
    labels.lineWidth,
    String(resolved.line?.width ?? 1.5),
    (value) => {
      const width = Number(value);
      currentLine = patchLine(currentLine, {
        width: Number.isFinite(width) ? Math.max(0, Math.min(12, width)) : 1.5,
      });
      styleTarget(draft, seriesId, pointId, {
        line: currentLine,
      });
    },
    'number',
  );
  makeSelect(
    body,
    labels.lineDash,
    resolved.line?.dash ?? 'solid',
    [
      { value: 'solid', label: labels.dashSolid },
      { value: 'dash', label: labels.dashDash },
      { value: 'dot', label: labels.dashDot },
      { value: 'dash-dot', label: labels.dashDashDot },
    ],
    (value) => {
      currentLine = patchLine(currentLine, { dash: value as ChartLineSpec['dash'] });
      styleTarget(draft, seriesId, pointId, { line: currentLine });
    },
  );
  // Smoothing belongs to a series' line, so it shows where a line or scatter series draws one.
  const drawsLine =
    chart.kind === 'line' ||
    (chart.kind === 'scatter' && (chart.scatterStyle ?? 'marker') !== 'marker');
  if (pointId === null && drawsLine) {
    makeCheckbox(body, labels.smoothLine, style.smooth === true, (checked) => {
      updateSeriesStyle(draft, seriesId, { smooth: checked });
    });
  }
};

const renderMarkerControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
): void => {
  const style = styleForSeries(draft, seriesId);
  const point =
    pointId === null ? undefined : style.points?.find((entry) => entry.pointId === pointId);
  const resolved = resolvedStyleForSelection(chart, data, draft, seriesId, pointId);
  let marker = { ...(pointId === null ? (style.marker ?? {}) : (point?.marker ?? {})) };
  const patchMarker = (patch: Partial<ChartMarkerSpec>): void => {
    marker = { ...marker, ...patch };
    styleTarget(draft, seriesId, pointId, { marker });
  };
  makeCheckbox(
    body,
    `${labels.marker} ${labels.markerVisible}`,
    resolved.marker.visible !== false,
    (checked) => {
      patchMarker({ visible: checked });
    },
  );
  makeSelect(
    body,
    labels.markerShape,
    resolved.marker.shape ?? 'circle',
    [
      { value: 'circle', label: labels.shapeCircle },
      { value: 'square', label: labels.shapeSquare },
      { value: 'diamond', label: labels.shapeDiamond },
      { value: 'triangle', label: labels.shapeTriangle },
      { value: 'x', label: labels.shapeX },
      { value: 'star', label: labels.shapeStar },
      { value: 'plus', label: labels.shapePlus },
      { value: 'dot', label: labels.shapeDot },
      { value: 'dash', label: labels.shapeDash },
    ],
    (value) => patchMarker({ shape: value as ChartMarkerSpec['shape'] }),
  );
  makeTextInput(
    body,
    labels.markerSize,
    String(resolved.marker.size ?? 5),
    (value) => {
      const size = Number(value);
      patchMarker({ size: Number.isFinite(size) ? Math.max(2, Math.min(24, size)) : 6 });
    },
    'number',
  );
  makeColorField(
    body,
    `${labels.marker} ${labels.fillColor}`,
    resolved.marker.fill?.color ?? resolved.fill?.color ?? resolved.color,
    (color) => {
      patchMarker({ fill: patchFill(marker.fill, { color }) });
    },
  );
  makeColorField(
    body,
    `${labels.marker} ${labels.lineColor}`,
    resolved.marker.line?.color ?? resolved.line?.color ?? resolved.color,
    (color) => {
      patchMarker({ line: patchLine(marker.line, { color }) });
    },
  );
};

const renderChartGeometryControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  draft: ChartDraft,
): void => {
  if (chart.kind === 'bar' || chart.kind === 'column') {
    const geometry = element('fieldset', 'fc-chart-settings__group');
    geometry.appendChild(textNode('legend', labels.gapWidth, 'fc-chart-settings__legend'));
    makeTextInput(
      geometry,
      labels.gapWidth,
      String(draft.gapWidth),
      (value) => {
        const parsed = Number(value);
        draft.gapWidth = Number.isFinite(parsed) ? Math.max(0, Math.min(500, parsed)) : 150;
      },
      'number',
      { min: 0, max: 500, step: '1' },
    );
    makeTextInput(
      geometry,
      labels.overlap,
      String(draft.overlap),
      (value) => {
        const parsed = Number(value);
        draft.overlap = Number.isFinite(parsed) ? Math.max(-100, Math.min(100, parsed)) : 0;
      },
      'number',
      {
        allowNegative: true,
        min: -100,
        max: 100,
        step: '1',
        disabled: draft.grouping !== 'clustered',
      },
    );
    body.appendChild(geometry);
  }
  if (isCircularKind(chart.kind)) {
    const geometry = element('fieldset', 'fc-chart-settings__group');
    geometry.appendChild(textNode('legend', labels.firstSliceAngle, 'fc-chart-settings__legend'));
    makeTextInput(
      geometry,
      labels.firstSliceAngle,
      String(draft.firstSliceAngle),
      (value) => {
        const parsed = Number(value);
        draft.firstSliceAngle = Number.isFinite(parsed) ? Math.max(0, Math.min(360, parsed)) : 0;
      },
      'number',
      { min: 0, max: 360, step: '1' },
    );
    body.appendChild(geometry);
  }
  if (isDoughnutKind(chart.kind)) {
    const geometry = element('fieldset', 'fc-chart-settings__group');
    geometry.appendChild(textNode('legend', labels.holeSize, 'fc-chart-settings__legend'));
    makeTextInput(
      geometry,
      labels.holeSize,
      String(draft.holeSize),
      (value) => {
        const parsed = Number(value);
        draft.holeSize = Number.isFinite(parsed) ? Math.max(10, Math.min(90, parsed)) : 50;
        draft.holeSizeChanged = draft.holeSize !== draft.holeSizeBaseline;
      },
      'number',
      { min: 10, max: 90, step: '1' },
    );
    body.appendChild(geometry);
  }
};

type NewSeriesField = keyof SessionChartSeriesRefs;

/** The add-series form as typed, kept across re-renders. */
interface NewSeriesDraft {
  name: string;
  values: string;
  categories: string;
  x: string;
  invalid: NewSeriesField[];
}

interface StylesState {
  seriesId: string;
  pointId: string | null;
  newSeries: NewSeriesDraft | null;
}

/** Series add and remove, which change the workbook chart at once rather than the draft. */
interface SeriesActions {
  /** The workbook chart's kind, which decides the references a new series takes. */
  readonly kind: ChartKind;
  /** Why both actions are unavailable; null when they are available. */
  readonly locked: string | null;
  readonly disabled: boolean;
  add(refs: SessionChartSeriesRefs): string | null;
  remove(seriesId: string): boolean;
}

/** The typed references, or the fields that are not references. Only `values` is required. */
const newSeriesRefs = (
  form: NewSeriesDraft,
  scatter: boolean,
): SessionChartSeriesRefs | NewSeriesField[] => {
  const fields: readonly NewSeriesField[] = ['name', 'values', scatter ? 'x' : 'categories'];
  const invalid = fields.filter((field) => {
    const text = form[field].trim();
    return text ? parseRangeRef(text) === null : field === 'values';
  });
  if (invalid.length > 0) return invalid;
  const refs: SessionChartSeriesRefs = { values: form.values.trim() };
  for (const field of fields) {
    const text = form[field].trim();
    if (field !== 'values' && text) refs[field] = text;
  }
  return refs;
};

const renderSeriesActions = (
  target: HTMLElement,
  labels: SessionChartSettingsLabels,
  seriesCount: number,
  state: StylesState,
  rerender: () => void,
  actions: SeriesActions,
): void => {
  const row = element('div', 'fc-chart-settings__series-actions');
  const add = element('button', 'fc-chart-settings__add-series');
  add.type = 'button';
  add.textContent = labels.addSeries;
  add.setAttribute('aria-expanded', String(state.newSeries !== null));
  add.addEventListener('click', () => {
    state.newSeries = state.newSeries
      ? null
      : { name: '', values: '', categories: '', x: '', invalid: [] };
    rerender();
  });
  const remove = element('button', 'fc-chart-settings__remove-series');
  remove.type = 'button';
  remove.textContent = labels.removeSeries;
  remove.addEventListener('click', () => {
    if (actions.remove(state.seriesId)) rerender();
  });
  row.append(add, remove);
  target.appendChild(row);
  const unavailable = actions.locked !== null || actions.disabled;
  projectDisabledState(add, unavailable, actions.locked);
  projectDisabledState(
    remove,
    unavailable || seriesCount < 2,
    actions.locked ?? (seriesCount < 2 ? labels.lastSeriesLocked : null),
  );
  const form = state.newSeries;
  if (!form || unavailable) return;

  const scatter = actions.kind === 'scatter';
  const group = element('div', 'fc-chart-settings__new-series');
  const fields: readonly [NewSeriesField, string][] = [
    ['name', labels.seriesNameRef],
    ['values', scatter ? labels.seriesYValuesRef : labels.seriesValuesRef],
    scatter ? ['x', labels.seriesXValuesRef] : ['categories', labels.axisLabelsRef],
  ];
  for (const [field, label] of fields) {
    const input = makeTextInput(group, label, form[field], (value) => {
      form[field] = value;
    });
    input.dataset.seriesRef = field;
    if (form.invalid.includes(field)) input.setAttribute('aria-invalid', 'true');
  }
  if (form.invalid.length > 0) {
    const error = textNode('div', labels.invalidReference, 'fc-chart-settings__error');
    error.setAttribute('role', 'alert');
    group.appendChild(error);
  }
  const buttons = element('div', 'fc-chart-settings__new-series-actions');
  const confirm = element('button', 'fc-chart-settings__new-series-confirm');
  confirm.type = 'button';
  confirm.textContent = labels.addSeriesConfirm;
  confirm.addEventListener('click', () => {
    const refs = newSeriesRefs(form, scatter);
    if (Array.isArray(refs)) {
      form.invalid = refs;
      rerender();
      return;
    }
    const added = actions.add(refs);
    if (added === null) {
      // The engine resolves what the syntax check cannot, such as an unknown sheet.
      form.invalid = ['values'];
    } else {
      state.seriesId = added;
      state.pointId = null;
      state.newSeries = null;
    }
    rerender();
  });
  const dismiss = element('button', 'fc-chart-settings__new-series-cancel');
  dismiss.type = 'button';
  dismiss.textContent = labels.cancel;
  dismiss.addEventListener('click', () => {
    state.newSeries = null;
    rerender();
  });
  buttons.append(confirm, dismiss);
  group.appendChild(buttons);
  target.appendChild(group);
};

const renderStyles = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  state: StylesState,
  rerender: () => void,
  accents: readonly string[] | undefined,
  seriesActions: SeriesActions,
): void => {
  const automatic = automaticPalette(accents);
  const presets = { ...PALETTE_PRESETS, office: automatic.slice(0, 6) };
  const palette = element('fieldset', 'fc-chart-settings__group');
  palette.appendChild(textNode('legend', labels.palette, 'fc-chart-settings__legend'));
  const presetRow = element('div', 'fc-chart-settings__palette-presets');
  for (const [id, colors] of Object.entries(presets)) {
    const button = element('button');
    button.type = 'button';
    button.className = 'fc-chart-settings__palette-button';
    button.dataset.palette = id;
    const paletteLabel = PALETTE_LABEL_KEYS[id];
    button.setAttribute('aria-label', paletteLabel ? labels[paletteLabel] : labels.palette);
    const swatches = element('span', 'fc-chart-settings__palette-swatches');
    for (const color of colors) {
      const swatch = element('span', 'fc-chart-settings__palette-swatch');
      swatch.style.backgroundColor = color;
      swatches.appendChild(swatch);
    }
    button.appendChild(swatches);
    button.addEventListener('click', () => {
      draft.palette = [...colors];
      releaseSeriesColors(draft);
      rerender();
    });
    presetRow.appendChild(button);
  }
  palette.appendChild(presetRow);
  for (let index = 0; index < Math.max(6, draft.palette.length); index += 1) {
    const color =
      draft.palette[index] ?? automatic[index % automatic.length] ?? DEFAULT_THEME_ACCENTS[0] ?? '';
    makeColorField(palette, `${labels.palette} ${index + 1}`, color, (next) => {
      draft.palette[index] = next;
      releaseSeriesColors(draft);
    });
  }
  makeCheckbox(palette, labels.varyColors, draft.varyColors, (checked) => {
    draft.varyColors = checked;
    draft.varyColorsManuallyChanged = true;
    draft.varyColorsOverride =
      checked === (draft.varyColorsOriginalOverride ?? defaultVaryColors(String(draft.kind)))
        ? draft.varyColorsOriginalOverride
        : checked;
  });
  body.appendChild(palette);

  const target = element('fieldset', 'fc-chart-settings__group');
  target.appendChild(textNode('legend', labels.series, 'fc-chart-settings__legend'));
  const entries = seriesEntries(data, labels);
  const selectedSeries = entries.some((entry) => entry.id === state.seriesId)
    ? state.seriesId
    : (entries[0]?.id ?? 'series-0');
  state.seriesId = selectedSeries;
  const points = stylePointEntries(data, selectedSeries, labels);
  if (state.pointId !== null && !points.some((point) => point.id === state.pointId)) {
    state.pointId = null;
  }
  makeSelect(
    target,
    labels.series,
    selectedSeries,
    entries.map((entry) => ({ value: entry.id, label: entry.label })),
    (value) => {
      state.seriesId = value;
      state.pointId = null;
      rerender();
    },
  );
  renderSeriesActions(target, labels, entries.length, state, rerender, seriesActions);
  makeSelect(
    target,
    labels.point,
    state.pointId ?? ALL_POINTS,
    [
      { value: ALL_POINTS, label: labels.allPoints },
      ...points.map((point) => ({ value: point.id, label: point.label })),
    ],
    (value) => {
      state.pointId = value === ALL_POINTS ? null : value;
      rerender();
    },
  );
  renderColorAndLineControls(target, labels, chart, data, draft, selectedSeries, state.pointId);
  renderMarkerControls(target, labels, chart, data, draft, selectedSeries, state.pointId);
  body.appendChild(target);
  renderChartGeometryControls(body, labels, chart, draft);
};

const renderAxisControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  axisName: AxisName,
  axisLabel: string,
  rerender: () => void,
): void => {
  const axis = draft[axisName];
  const enabled = !isCircularKind(chart.kind);
  const numeric = axisIsNumeric(chart, axisName);
  const logEligible = axisLogEligible(chart, axisName);
  const categoryCount = categoryCountForDraft(data, draft);
  const group = element('fieldset', 'fc-chart-settings__group');
  group.appendChild(textNode('legend', axisLabel, 'fc-chart-settings__legend'));
  makeCheckbox(
    group,
    labels.showAxis,
    axis.visible,
    (checked) => updateAxisField(draft, axisName, 'visible', checked),
    !enabled,
  );
  makeCheckbox(
    group,
    labels.reverseOrder,
    axis.reverseOrder,
    (checked) => updateAxisField(draft, axisName, 'reverseOrder', checked),
    !enabled,
  );
  makeCheckbox(
    group,
    labels.logarithmicScale,
    axis.logEnabled,
    (checked) => {
      updateAxisLogEnabled(draft, axisName, checked);
      rerender();
    },
    !enabled || !logEligible,
    'fc-chart-settings__log-scale',
  );
  const logBaseInput = makeTextInput(
    group,
    labels.logBase,
    String(axis.logBaseDraft),
    (value) => updateAxisLogBase(draft, axisName, value),
    'number',
    {
      min: 2,
      max: 1000,
      step: 'any',
      disabled: !enabled || !logEligible || !axis.logEnabled,
    },
  );
  logBaseInput.parentElement?.classList.add('fc-chart-settings__log-base');
  const crossingSelect = makeSelect(
    group,
    axisName === 'xAxis' ? labels.verticalAxisCrosses : labels.horizontalAxisCrosses,
    axis.crossingMode,
    [
      { value: 'auto', label: labels.automatic },
      { value: 'min', label: labels.minimum },
      { value: 'max', label: labels.maximum },
      { value: 'custom', label: labels.custom },
    ],
    (value) => {
      updateAxisCrossingMode(draft, axisName, value as AxisCrossingMode, numeric, categoryCount);
      rerender();
    },
    !enabled,
  );
  // Identify the mode row separately because its options reuse the bound
  // labels "Minimum" and "Maximum".
  crossingSelect.parentElement?.classList.add('fc-chart-settings__crossing-mode');
  makeTextInput(
    group,
    labels.customValue,
    String(normalizedCrossingCustom(axis.crossingCustom, numeric, categoryCount)),
    (value) => updateAxisCrossingCustom(draft, axisName, value, numeric, categoryCount),
    'number',
    {
      allowNegative: numeric,
      min: numeric ? undefined : 1,
      step: numeric ? 'any' : '1',
      disabled: !enabled || axis.crossingMode !== 'custom',
    },
  );
  const labelPositionSelect = makeSelect(
    group,
    labels.labelPosition,
    axis.labelPositionMode,
    [
      { value: 'automatic', label: labels.automatic },
      { value: 'nextTo', label: labels.labelPositionNextTo },
      { value: 'low', label: labels.labelPositionLow },
      { value: 'high', label: labels.labelPositionHigh },
      { value: 'none', label: labels.labelPositionNone },
    ],
    (value) => {
      updateAxisLabelPosition(draft, axisName, value as AxisLabelPositionMode);
    },
    !enabled,
  );
  labelPositionSelect.parentElement?.classList.add('fc-chart-settings__label-position');
  // Typing a title replaces the automatic one; the checkbox returns to it.
  let automatic: HTMLInputElement | undefined;
  makeTextInput(
    group,
    axisName === 'xAxis' ? labels.xAxisTitle : labels.yAxisTitle,
    axis.title,
    (value) => {
      updateAxisField(draft, axisName, 'title', value);
      if (value.trim().length > 0 && draft[axisName].titleAuto) {
        updateAxisField(draft, axisName, 'titleAuto', false);
        if (automatic) automatic.checked = false;
      }
    },
    'text',
    { disabled: !enabled },
  );
  automatic = makeCheckbox(
    group,
    labels.titleAutomatic,
    axis.titleAuto,
    (checked) => updateAxisField(draft, axisName, 'titleAuto', checked),
    !enabled,
  );
  makeTextInput(
    group,
    labels.minimum,
    axis.min === undefined ? '' : String(axis.min),
    (value) => updateAxisField(draft, axisName, 'min', parseOptionalNumber(value)),
    'number',
    {
      allowNegative: true,
      step: 'any',
      placeholder: labels.automatic,
      disabled: !enabled || !numeric,
    },
  );
  makeTextInput(
    group,
    labels.maximum,
    axis.max === undefined ? '' : String(axis.max),
    (value) => updateAxisField(draft, axisName, 'max', parseOptionalNumber(value)),
    'number',
    {
      allowNegative: true,
      step: 'any',
      placeholder: labels.automatic,
      disabled: !enabled || !numeric,
    },
  );
  makeTextInput(
    group,
    labels.majorUnit,
    axis.majorUnit === undefined ? '' : String(axis.majorUnit),
    (value) => {
      const parsed = parseOptionalNumber(value);
      updateAxisField(
        draft,
        axisName,
        'majorUnit',
        parsed !== undefined && parsed > 0 ? parsed : undefined,
      );
    },
    'number',
    {
      min: 0,
      step: 'any',
      placeholder: labels.automatic,
      disabled: !enabled || !numeric || (logEligible && axis.logEnabled),
    },
  );
  makeTextInput(
    group,
    labels.axisNumberFormat,
    axis.numberFormat,
    (value) => updateAxisField(draft, axisName, 'numberFormat', value),
    'text',
    { disabled: !enabled || !numeric },
  );
  makeCheckbox(
    group,
    labels.majorGridlines,
    axis.majorGridlines,
    (checked) => updateAxisField(draft, axisName, 'majorGridlines', checked),
    !enabled || !numeric,
  );
  body.appendChild(group);
};

const renderElements = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  rerender: () => void,
): void => {
  const general = element('fieldset', 'fc-chart-settings__group');
  general.appendChild(textNode('legend', labels.title, 'fc-chart-settings__legend'));
  // Typing a title replaces the automatic one; the checkbox returns to it.
  let automatic: HTMLInputElement | undefined;
  makeTextInput(general, labels.title, draft.title, (value) => {
    draft.title = value;
    if (value.trim().length > 0 && draft.titleAuto) {
      draft.titleAuto = false;
      if (automatic) automatic.checked = false;
    }
  });
  automatic = makeCheckbox(general, labels.titleAutomatic, draft.titleAuto, (checked) => {
    draft.titleAuto = checked;
  });
  makeSelect(
    general,
    labels.chartType,
    String(draft.kind),
    [
      { value: 'column', label: labels.columnChart },
      { value: 'bar', label: labels.barChart },
      { value: 'line', label: labels.lineChart },
      { value: 'area', label: labels.areaChart },
      { value: 'pie', label: labels.pieChart },
      { value: 'doughnut', label: labels.doughnutChart },
      { value: 'scatter', label: labels.scatterChart },
    ],
    (value) => {
      draft.kind = value as ChartDraft['kind'];
      draft.kindDirty = draft.kind !== draft.kindBaseline;
      draft.kindChanged = draft.kindDirty || draft.seriesDirectionDirty;
      draft.legendKindChanged = draft.kindChanged;
      draft.varyColorsKindChanged = draft.kindChanged;
      rerender();
    },
  );
  if (chart.kind === 'scatter') {
    makeSelect(
      general,
      labels.scatterStyle,
      draft.scatterStyle,
      [
        { value: 'marker', label: labels.scatterStyleMarker },
        { value: 'line', label: labels.scatterStyleLine },
        { value: 'lineMarker', label: labels.scatterStyleLineMarker },
      ],
      (value) => {
        const next = value as ChartScatterStyle;
        draft.scatterStyle = next;
        // Like Excel's subtype change, the style rewrites any explicit marker visibility.
        draft.seriesStyles = draft.seriesStyles.map((style) =>
          style.marker?.visible === undefined
            ? style
            : { ...style, marker: { ...style.marker, visible: next !== 'line' } },
        );
        draft.scatterStyleChanged = next !== draft.scatterStyleBaseline;
        draft.scatterStyleOverride =
          next === draft.scatterStyleBaseline ? draft.scatterStyleOriginalOverride : next;
        rerender();
      },
    );
  }
  makeCheckbox(general, labels.legend, draft.legendVisible, (checked) => {
    draft.legendVisible = checked;
    draft.legendManuallyChanged = true;
    draft.legendOverride =
      checked === (draft.legendOriginalOverride ?? effectiveLegendDefault(chart, draft, data))
        ? draft.legendOriginalOverride
        : checked;
  });
  makeSelect(
    general,
    labels.legendPosition,
    draft.legendPosition,
    [
      { value: 'top', label: labels.positionTop },
      { value: 'right', label: labels.positionRight },
      { value: 'bottom', label: labels.positionBottom },
      { value: 'left', label: labels.positionLeft },
    ],
    (value) => {
      draft.legendPosition = value as ChartDraft['legendPosition'];
    },
  );
  makeSelect(
    general,
    labels.grouping,
    draft.grouping,
    [
      { value: 'clustered', label: labels.groupingClustered },
      { value: 'stacked', label: labels.groupingStacked },
      { value: 'percent-stacked', label: labels.groupingPercentStacked },
    ],
    (value) => {
      draft.grouping = value as ChartGrouping;
      rerender();
    },
    !groupingSupported(chart),
  );
  const direction = makeSelect(
    general,
    labels.seriesDirection,
    draft.seriesDirection,
    [
      { value: 'auto', label: labels.seriesDirectionAuto },
      { value: 'columns', label: labels.seriesDirectionColumns },
      { value: 'rows', label: labels.seriesDirectionRows },
    ],
    (value) => {
      draft.seriesDirection = value as ChartSeriesDirection;
      draft.seriesDirectionDirty = draft.seriesDirection !== draft.seriesDirectionBaseline;
      draft.kindChanged = draft.kindDirty || draft.seriesDirectionDirty;
      draft.legendKindChanged = draft.kindChanged;
      draft.varyColorsKindChanged = draft.kindChanged;
      rerender();
    },
  );
  // A PivotChart's series follow its PivotTable.
  if (chart.pivotChart) projectDisabledState(direction, true, labels.pivotChartLocked);
  makeSelect(
    general,
    labels.emptyCells,
    emptyCellsForcedZero(chart, draft.grouping) ? 'zero' : draft.emptyCells,
    [
      { value: 'gap', label: labels.emptyCellsGap },
      { value: 'zero', label: labels.emptyCellsZero },
      {
        value: 'span',
        label: labels.emptyCellsSpan,
        disabled: emptyCellOptionDisabled(chart, draft.grouping, 'span'),
      },
    ],
    (value) => {
      draft.emptyCells = value as ChartEmptyCellMode;
    },
    emptyCellsForcedZero(chart, draft.grouping),
  );
  body.appendChild(general);

  const axes = element('fieldset', 'fc-chart-settings__group');
  axes.appendChild(textNode('legend', labels.axisTitles, 'fc-chart-settings__legend'));
  renderAxisControls(axes, labels, chart, data, draft, 'xAxis', labels.xAxis, rerender);
  renderAxisControls(axes, labels, chart, data, draft, 'yAxis', labels.yAxis, rerender);
  body.appendChild(axes);

  const dataLabels = element('fieldset', 'fc-chart-settings__group');
  dataLabels.appendChild(textNode('legend', labels.dataLabels, 'fc-chart-settings__legend'));
  const value = draft.dataLabels;
  makeCheckbox(dataLabels, labels.showDataLabels, value.visible, (checked) => {
    updateDataLabel(draft, 'visible', checked);
  });
  makeCheckbox(dataLabels, labels.showSeriesName, value.showSeriesName, (checked) => {
    updateDataLabel(draft, 'showSeriesName', checked);
  });
  makeCheckbox(dataLabels, labels.showCategoryName, value.showCategoryName, (checked) => {
    updateDataLabel(draft, 'showCategoryName', checked);
  });
  makeCheckbox(dataLabels, labels.showValue, value.showValue, (checked) => {
    updateDataLabel(draft, 'showValue', checked);
  });
  makeCheckbox(
    dataLabels,
    labels.showPercentage,
    value.showPercentage,
    (checked) => {
      updateDataLabel(draft, 'showPercentage', checked);
    },
    !supportsPercentage(chart, draft.grouping),
  );
  makeCheckbox(dataLabels, labels.showLegendKey, value.showLegendKey, (checked) => {
    updateDataLabel(draft, 'showLegendKey', checked);
  });
  makeTextInput(dataLabels, labels.labelSeparator, value.separator, (next) => {
    updateDataLabel(draft, 'separator', next);
  });
  makeSelect(
    dataLabels,
    labels.labelPosition,
    value.position,
    allowedPositions({ ...chart, grouping: draft.grouping }, value.position).map((item) => ({
      value: item.value,
      label: labels[item.label],
      disabled: item.disabled,
    })),
    (next) => {
      updateDataLabel(draft, 'position', next as MutableDataLabels['position']);
    },
    chart.kind === 'area' || isDoughnutKind(chart.kind),
  );
  makeTextInput(dataLabels, labels.numberFormat, value.numberFormat, (next) => {
    updateDataLabel(draft, 'numberFormat', next);
  });
  makeColorField(dataLabels, labels.textColor, value.color, (next) => {
    updateDataLabel(draft, 'color', next);
  });
  makeTextInput(
    dataLabels,
    labels.fontSize,
    String(value.fontSize),
    (next) => {
      const size = Number(next);
      updateDataLabel(
        draft,
        'fontSize',
        Number.isFinite(size) ? Math.max(6, Math.min(72, size)) : 10,
      );
    },
    'number',
  );
  body.appendChild(dataLabels);
};

const renderFilters = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  data: ChartData,
  draft: ChartDraft,
): void => {
  const series = element('fieldset', 'fc-chart-settings__group');
  series.appendChild(textNode('legend', labels.series, 'fc-chart-settings__legend'));
  const seriesEntriesValue = seriesEntries(data, labels);
  if (seriesEntriesValue.length === 0) {
    series.appendChild(textNode('div', labels.noSeries, 'fc-chart-settings__empty'));
  } else {
    const actions = element('div', 'fc-chart-settings__filter-actions');
    const all = element('button');
    all.type = 'button';
    all.textContent = labels.selectAll;
    all.addEventListener('click', () => {
      const known = new Set(seriesEntriesValue.map((entry) => entry.id));
      draft.filter.hiddenSeriesIds = draft.filter.hiddenSeriesIds.filter((id) => !known.has(id));
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    const clear = element('button');
    clear.type = 'button';
    clear.textContent = labels.clearAll;
    clear.addEventListener('click', () => {
      draft.filter.hiddenSeriesIds = [
        ...new Set([
          ...draft.filter.hiddenSeriesIds,
          ...seriesEntriesValue.map((entry) => entry.id),
        ]),
      ];
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    actions.append(all, clear);
    series.appendChild(actions);
    for (const entry of seriesEntriesValue) {
      makeCheckbox(
        series,
        entry.label,
        !draft.filter.hiddenSeriesIds.includes(entry.id),
        (checked) => {
          draft.filter.hiddenSeriesIds = checked
            ? draft.filter.hiddenSeriesIds.filter((id) => id !== entry.id)
            : [...new Set([...draft.filter.hiddenSeriesIds, entry.id])];
        },
        false,
        'fc-chart-settings__filter-item',
      );
    }
  }
  body.appendChild(series);

  const categories = element('fieldset', 'fc-chart-settings__group');
  categories.appendChild(textNode('legend', labels.categories, 'fc-chart-settings__legend'));
  const categoryEntriesValue = categoryEntries(data, labels);
  if (categoryEntriesValue.length === 0) {
    categories.appendChild(textNode('div', labels.noCategories, 'fc-chart-settings__empty'));
  } else {
    const actions = element('div', 'fc-chart-settings__filter-actions');
    const all = element('button');
    all.type = 'button';
    all.textContent = labels.selectAll;
    all.addEventListener('click', () => {
      const known = new Set(categoryEntriesValue.map((entry) => entry.id));
      draft.filter.hiddenCategoryIds = draft.filter.hiddenCategoryIds.filter(
        (id) => !known.has(id),
      );
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    const clear = element('button');
    clear.type = 'button';
    clear.textContent = labels.clearAll;
    clear.addEventListener('click', () => {
      draft.filter.hiddenCategoryIds = [
        ...new Set([
          ...draft.filter.hiddenCategoryIds,
          ...categoryEntriesValue.map((entry) => entry.id),
        ]),
      ];
      body.replaceChildren();
      renderFilters(body, labels, data, draft);
    });
    actions.append(all, clear);
    categories.appendChild(actions);
    for (const entry of categoryEntriesValue) {
      makeCheckbox(
        categories,
        entry.label,
        !draft.filter.hiddenCategoryIds.includes(entry.id),
        (checked) => {
          draft.filter.hiddenCategoryIds = checked
            ? draft.filter.hiddenCategoryIds.filter((id) => id !== entry.id)
            : [...new Set([...draft.filter.hiddenCategoryIds, entry.id])];
        },
        false,
        'fc-chart-settings__filter-item',
      );
    }
  }
  body.appendChild(categories);
};

/**
 * Create the non-modal settings surface used by session chart overlays. The
 * surface is intentionally independent from the chart root: chart redraws can
 * replace SVG panels without destroying an open form or its draft.
 */
export function createSessionChartSettings(
  deps: SessionChartSettingsDeps,
): SessionChartSettingsHandle {
  const { host, history = null } = deps;
  let labels = labelsWithDefaults(deps.labels);
  let locale = deps.locale ?? 'en-US';
  let chart: SessionChartSettingsShape | null = null;
  let data: ChartData | null = null;
  let anchor: HTMLElement | null = null;
  let draft: ChartDraft | null = null;
  let activeTab: SettingsTab = 'elements';
  let styleSeriesId = 'series-0';
  let stylePointId: string | null = null;
  let newSeries: NewSeriesDraft | null = null;
  let dataKey = '';
  let editable = true;
  let disposed = false;

  const surface = element('div', 'fc-chart-settings');
  surface.hidden = true;
  surface.setAttribute('role', 'dialog');
  surface.setAttribute('aria-modal', 'false');
  surface.tabIndex = -1;
  host.appendChild(surface);

  const removeEscape = (): void => {
    document.removeEventListener('keydown', onDocumentKeydown);
  };

  const closeInternal = (returnFocus: boolean): void => {
    if (surface.hidden) return;
    const returnTarget = anchor;
    surface.hidden = true;
    removeEscape();
    chart = null;
    data = null;
    anchor = null;
    draft = null;
    newSeries = null;
    dataKey = '';
    if (returnFocus && returnTarget?.isConnected) returnTarget.focus({ preventScroll: true });
  };

  const reposition = (): void => {
    if (!anchor) return;
    const hostRect = host.getBoundingClientRect();
    const anchorRect = anchor.getBoundingClientRect();
    const hostScrollLeft = host.scrollLeft;
    const hostScrollTop = host.scrollTop;
    const hostWidth = host.clientWidth || hostRect.width || 320;
    const hostHeight = host.clientHeight || hostRect.height || 520;
    const surfaceWidth = surface.offsetWidth || 320;
    const surfaceHeight = surface.offsetHeight || 360;
    const anchorWidth = anchorRect.width || anchor.offsetWidth || 0;
    const anchorHeight = anchorRect.height || anchor.offsetHeight || 0;
    const hasAnchorRect =
      anchorRect.width !== 0 ||
      anchorRect.height !== 0 ||
      anchorRect.left !== 0 ||
      anchorRect.top !== 0 ||
      anchorRect.right !== 0 ||
      anchorRect.bottom !== 0;
    const anchorLeft = hasAnchorRect
      ? anchorRect.left - hostRect.left + hostScrollLeft
      : anchor.offsetLeft || 0;
    const anchorTop = hasAnchorRect
      ? anchorRect.top - hostRect.top + hostScrollTop
      : anchor.offsetTop || 0;
    const anchorRight = anchorLeft + anchorWidth;
    const anchorBottom = anchorTop + anchorHeight;
    const minLeft = hostScrollLeft + 4;
    const maxLeft = Math.max(minLeft, hostScrollLeft + hostWidth - surfaceWidth - 4);
    const minTop = hostScrollTop + 4;
    const maxTop = Math.max(minTop, hostScrollTop + hostHeight - surfaceHeight - 4);
    let left = anchorRight + 8;
    if (left > maxLeft) left = anchorLeft - surfaceWidth - 8;
    let top = anchorTop;
    if (top > maxTop) top = anchorBottom - surfaceHeight;
    surface.style.left = `${Math.max(minLeft, Math.min(left, maxLeft))}px`;
    surface.style.top = `${Math.max(minTop, Math.min(top, maxTop))}px`;
  };

  const syncUntouchedOptionalOverrides = (): void => {
    if (!chart || !data || !draft) return;
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

  const refreshDraftDefaults = (): void => {
    if (!chart || !data || !draft) return;
    const effective = effectiveDraftChart(chart, draft);
    if (draft.legendOverride === undefined) {
      draft.legendVisible = effectiveLegendVisible(effective, data);
    }
    if (draft.varyColorsOverride === undefined) {
      draft.varyColors = defaultVaryColors(String(draft.kind));
    }
  };

  const refreshDraftData = (): void => {
    if (!chart || !draft) return;
    if (deps.readData) data = deps.readData(effectiveDraftChart(chart, draft));
    refreshDraftDefaults();
  };

  const reconcileStyleSelection = (): void => {
    if (!data) return;
    const entries = seriesEntries(data);
    if (!entries.some((entry) => entry.id === styleSeriesId)) {
      styleSeriesId = entries[0]?.id ?? 'series-0';
    }
    if (stylePointId !== null) {
      const points = stylePointEntries(data, styleSeriesId);
      if (!points.some((point) => point.id === stylePointId)) stylePointId = null;
    }
  };

  /** Align the draft with series added or removed in the workbook. */
  const reconcileSeriesDraft = (): void => {
    if (!chart || !draft) return;
    const known = new Set(chart.data.series.map((series) => series.id));
    const styles = draft.seriesStyles.filter((style) => known.has(style.seriesId));
    for (const style of cloneStyles(chart.seriesStyles) ?? []) {
      if (!styles.some((entry) => entry.seriesId === style.seriesId)) styles.push(style);
    }
    draft.seriesStyles = styles;
    draft.filter.hiddenSeriesIds = draft.filter.hiddenSeriesIds.filter((id) => known.has(id));
  };

  const renderPanel = (): void => {
    if (!chart || !data || !draft || disposed) return;
    refreshDraftDefaults();
    surface.replaceChildren();
    surface.setAttribute('aria-label', labels.settings);
    const heading = element('div', 'fc-chart-settings__header');
    const headingText = textNode('h3', labels.settings, 'fc-chart-settings__title');
    const close = element('button');
    close.type = 'button';
    close.className = 'fc-chart-settings__close';
    close.setAttribute('aria-label', labels.closeSettings);
    close.textContent = '×';
    close.addEventListener('click', () => closeInternal(true));
    heading.append(headingText, close);

    const tabs = element('div', 'fc-chart-settings__tabs');
    tabs.setAttribute('role', 'tablist');
    const tabDefinitions: readonly { id: SettingsTab; label: string }[] = [
      { id: 'elements', label: labels.elementsTab },
      { id: 'styles', label: labels.stylesTab },
      { id: 'filters', label: labels.filtersTab },
    ];
    for (const [index, item] of tabDefinitions.entries()) {
      const button = element('button');
      button.type = 'button';
      button.className = 'fc-chart-settings__tab';
      button.id = `fc-chart-settings-tab-${item.id}`;
      button.setAttribute('role', 'tab');
      button.setAttribute('aria-selected', String(activeTab === item.id));
      button.tabIndex = activeTab === item.id ? 0 : -1;
      button.textContent = item.label;
      button.addEventListener('click', () => {
        activeTab = item.id;
        renderPanel();
        surface.querySelector<HTMLButtonElement>(`#fc-chart-settings-tab-${item.id}`)?.focus();
      });
      button.addEventListener('keydown', (event) => {
        if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
        event.preventDefault();
        const nextIndex =
          (index + (event.key === 'ArrowRight' ? 1 : -1) + tabDefinitions.length) %
          tabDefinitions.length;
        const next = tabDefinitions[nextIndex];
        if (!next) return;
        activeTab = next.id;
        renderPanel();
        surface.querySelector<HTMLButtonElement>(`#fc-chart-settings-tab-${next.id}`)?.focus();
      });
      tabs.appendChild(button);
    }

    const body = element('div', 'fc-chart-settings__body');
    const renderAgain = (): void => {
      const view = captureSettingsView(surface);
      refreshDraftData();
      reconcileStyleSelection();
      renderPanel();
      restoreSettingsView(surface, view);
    };
    const effective = effectiveDraftChart(chart, draft);
    if (activeTab === 'elements') renderElements(body, labels, effective, data, draft, renderAgain);
    else if (activeTab === 'styles') {
      const state: StylesState = { seriesId: styleSeriesId, pointId: stylePointId, newSeries };
      const committed = chart;
      // The series of a draft that re-splits them are not the workbook's yet.
      const seriesActions: SeriesActions = {
        kind: committed.kind,
        locked: committed.pivotChart ? labels.pivotChartLocked : null,
        disabled: draft.kindChanged,
        add: (refs) => {
          const added = deps.onAddSeries(committed.id, refs, history);
          if (added !== null) reconcileSeriesDraft();
          return added;
        },
        remove: (seriesId) => {
          const removed = deps.onRemoveSeries(committed.id, seriesId, history);
          if (removed) reconcileSeriesDraft();
          return removed;
        },
      };
      renderStyles(
        body,
        labels,
        effective,
        data,
        draft,
        state,
        () => {
          styleSeriesId = state.seriesId;
          stylePointId = state.pointId;
          newSeries = state.newSeries;
          renderAgain();
        },
        deps.accents?.(),
        seriesActions,
      );
    } else {
      renderFilters(body, labels, data, draft);
      if (chart.pivotChart) {
        for (const control of Array.from(
          body.querySelectorAll<HTMLInputElement | HTMLButtonElement>('input, button'),
        )) {
          projectDisabledState(control, true, labels.pivotChartLocked);
        }
      }
    }

    const footer = element('div', 'fc-chart-settings__footer');
    const cancel = element('button');
    cancel.type = 'button';
    cancel.className = 'fc-chart-settings__cancel';
    cancel.textContent = labels.cancel;
    cancel.addEventListener('click', () => closeInternal(true));
    const apply = element('button');
    apply.type = 'button';
    apply.className = 'fc-chart-settings__apply';
    apply.textContent = labels.apply;
    apply.addEventListener('click', () => {
      if (!chart || !draft) return;
      const id = chart.id;
      const patch = toPatch(draft, chart);
      // The callback is deliberately invoked once per Apply click. The
      // command owns protection and history; this form only owns the draft.
      if (Object.keys(patch).length > 0) deps.onApply(id, patch, history);
      closeInternal(true);
    });
    footer.append(cancel, apply);
    surface.append(heading, tabs, body, footer);
    if (!editable) {
      for (const control of Array.from(
        body.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLButtonElement>(
          'input, select, button',
        ),
      )) {
        projectDisabledState(control, true, null);
      }
      projectDisabledState(apply, true, null);
    }
    reposition();
  };

  function onDocumentKeydown(event: KeyboardEvent): void {
    if (surface.hidden || event.key !== 'Escape') return;
    event.preventDefault();
    closeInternal(true);
  }

  const open = (
    nextChart: SessionChartSettingsShape,
    nextData: ChartData,
    nextAnchor: HTMLElement,
    tab?: SettingsTab,
  ): void => {
    if (disposed) return;
    chart = nextChart;
    data = nextData;
    anchor = nextAnchor;
    draft = draftFromChart(nextChart, nextData, automaticPalette(deps.accents?.()));
    activeTab = tab ?? 'elements';
    styleSeriesId = seriesEntries(nextData)[0]?.id ?? 'series-0';
    stylePointId = null;
    newSeries = null;
    dataKey = dataSignature(nextData);
    editable = deps.canEdit?.(nextChart) !== false;
    surface.hidden = false;
    surface.lang = locale;
    document.removeEventListener('keydown', onDocumentKeydown);
    document.addEventListener('keydown', onDocumentKeydown);
    renderPanel();
    surface
      .querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')
      ?.focus({ preventScroll: true });
  };

  return {
    open,
    sync(nextChart, nextData, nextAnchor) {
      if (surface.hidden || !chart || chart.id !== nextChart.id || !draft) return;
      const previousKind = draft.kind;
      const previousSeriesDirection = draft.seriesDirection;
      const previousLegendVisible = draft.legendVisible;
      const previousVaryColors = draft.varyColors;
      const previousScatterStyle = draft.scatterStyle;
      const previousXReverseOrder = draft.xAxis.reverseOrder;
      const previousYReverseOrder = draft.yAxis.reverseOrder;
      const previousXCrossing = draft.xAxis.crossing;
      const previousYCrossing = draft.yAxis.crossing;
      const previousXCrossingMode = draft.xAxis.crossingMode;
      const previousYCrossingMode = draft.yAxis.crossingMode;
      const previousXCrossingCustom = draft.xAxis.crossingCustom;
      const previousYCrossingCustom = draft.yAxis.crossingCustom;
      const previousXLabelPosition = draft.xAxis.labelPosition;
      const previousYLabelPosition = draft.yAxis.labelPosition;
      const previousXLabelPositionMode = draft.xAxis.labelPositionMode;
      const previousYLabelPositionMode = draft.yAxis.labelPositionMode;
      const previousXLogBase = draft.xAxis.logBase;
      const previousYLogBase = draft.yAxis.logBase;
      const previousXLogEnabled = draft.xAxis.logEnabled;
      const previousYLogEnabled = draft.yAxis.logEnabled;
      const previousXLogBaseDraft = draft.xAxis.logBaseDraft;
      const previousYLogBaseDraft = draft.yAxis.logBaseDraft;
      chart = nextChart;
      if (!draft.kindDirty) {
        draft.kind = nextChart.kind;
        draft.kindBaseline = nextChart.kind;
      }
      if (!draft.seriesDirectionDirty) {
        draft.seriesDirection = nextChart.seriesDirection ?? 'auto';
        draft.seriesDirectionBaseline = draft.seriesDirection;
      }
      if (deps.readData) data = deps.readData(effectiveDraftChart(nextChart, draft));
      else data = nextData;
      syncUntouchedOptionalOverrides();
      refreshDraftDefaults();
      reconcileStyleSelection();
      const nextEditable = deps.canEdit?.(nextChart) !== false;
      if (!anchor?.isConnected) {
        anchor =
          nextAnchor.querySelector<HTMLButtonElement>(`[data-chart-settings-tab="${activeTab}"]`) ??
          nextAnchor;
      }
      const nextKey = dataSignature(data);
      const permissionChanged = nextEditable !== editable;
      const chartFamilyChanged =
        draft.kind !== previousKind || draft.seriesDirection !== previousSeriesDirection;
      const optionalDefaultsChanged =
        draft.legendVisible !== previousLegendVisible || draft.varyColors !== previousVaryColors;
      const scatterStyleChangedExternally = draft.scatterStyle !== previousScatterStyle;
      const axisReverseOrderChanged =
        draft.xAxis.reverseOrder !== previousXReverseOrder ||
        draft.yAxis.reverseOrder !== previousYReverseOrder;
      const axisCrossingChanged =
        draft.xAxis.crossingMode !== previousXCrossingMode ||
        draft.yAxis.crossingMode !== previousYCrossingMode ||
        !Object.is(draft.xAxis.crossing, previousXCrossing) ||
        !Object.is(draft.yAxis.crossing, previousYCrossing) ||
        !Object.is(draft.xAxis.crossingCustom, previousXCrossingCustom) ||
        !Object.is(draft.yAxis.crossingCustom, previousYCrossingCustom);
      const axisLabelPositionChanged =
        draft.xAxis.labelPositionMode !== previousXLabelPositionMode ||
        draft.yAxis.labelPositionMode !== previousYLabelPositionMode ||
        !Object.is(draft.xAxis.labelPosition, previousXLabelPosition) ||
        !Object.is(draft.yAxis.labelPosition, previousYLabelPosition);
      const axisLogChanged =
        draft.xAxis.logEnabled !== previousXLogEnabled ||
        draft.yAxis.logEnabled !== previousYLogEnabled ||
        !Object.is(draft.xAxis.logBase, previousXLogBase) ||
        !Object.is(draft.yAxis.logBase, previousYLogBase) ||
        !Object.is(draft.xAxis.logBaseDraft, previousXLogBaseDraft) ||
        !Object.is(draft.yAxis.logBaseDraft, previousYLogBaseDraft);
      if (
        permissionChanged ||
        chartFamilyChanged ||
        optionalDefaultsChanged ||
        scatterStyleChangedExternally ||
        axisReverseOrderChanged ||
        axisCrossingChanged ||
        axisLabelPositionChanged ||
        axisLogChanged ||
        nextKey !== dataKey
      ) {
        const view = captureSettingsView(surface);
        editable = nextEditable;
        dataKey = nextKey;
        renderPanel();
        restoreSettingsView(surface, view);
      } else reposition();
    },
    close(discard = true) {
      closeInternal(discard);
    },
    setLabels(next, nextLocale) {
      labels = { ...labels, ...next };
      if (nextLocale) locale = nextLocale;
      if (!surface.hidden) {
        const view = captureSettingsView(surface);
        surface.lang = locale;
        renderPanel();
        restoreSettingsView(surface, view);
      }
    },
    detach() {
      if (disposed) return;
      disposed = true;
      removeEscape();
      surface.remove();
      chart = null;
      data = null;
      anchor = null;
      draft = null;
    },
    isOpen() {
      return !surface.hidden;
    },
    chartId() {
      return chart?.id ?? null;
    },
  };
}

export const chartSettingsDefaults = DEFAULT_LABELS;
export const chartSettingsPalettePresets = PALETTE_PRESETS;
