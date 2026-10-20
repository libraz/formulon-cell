import { seriesColorOf, withoutSeriesColor } from '../charts/engine-adapter.js';
import {
  accentsFromTheme,
  type ChartPalettePreset,
  chartPalettePreset,
} from '../charts/series-colors.js';
import type { ChartAxisSpec, ChartKind } from '../charts/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { type SessionChartPatch, sessionChartById, updateSessionChart } from './session-chart.js';

/** An entry of the Add Chart Element menu. Axes and gridlines are named by
 *  the direction they are drawn in, as the spreadsheet names them. */
export type ChartElementAction =
  | 'axis-title-horizontal'
  | 'axis-title-vertical'
  | 'title-none'
  | 'title-above'
  | 'labels-none'
  | 'labels-show'
  | 'gridlines-horizontal'
  | 'gridlines-vertical'
  | 'legend-none'
  | 'legend-right'
  | 'legend-top'
  | 'legend-left'
  | 'legend-bottom';

export const CHART_ELEMENT_ACTIONS: readonly ChartElementAction[] = [
  'axis-title-horizontal',
  'axis-title-vertical',
  'title-none',
  'title-above',
  'labels-none',
  'labels-show',
  'gridlines-horizontal',
  'gridlines-vertical',
  'legend-none',
  'legend-right',
  'legend-top',
  'legend-left',
  'legend-bottom',
];

type DrawableChart = SessionChart & { kind: ChartKind };

/** The chart selected on the active sheet, or null when none is selected. */
export const selectedSessionChart = (state: ReturnType<SpreadsheetStore['getState']>) => {
  const id = state.ui.selectedObject?.id;
  const chart = id === undefined ? null : sessionChartById(state, id);
  return chart && chart.sheet === state.data.sheetIndex ? chart : null;
};

const isDrawable = (chart: SessionChart): chart is DrawableChart => chart.kind !== 'unsupported';

const hasAxes = (kind: ChartKind): boolean => kind !== 'pie' && kind !== 'doughnut';

/** The spec key of the axis drawn horizontally (`'horizontal'`) or vertically. */
const axisKey = (direction: 'horizontal' | 'vertical'): 'xAxis' | 'yAxis' =>
  direction === 'horizontal' ? 'xAxis' : 'yAxis';

/** Whether the axis drawn in `direction` has values, the only axis with gridlines here. */
const isValueAxis = (kind: ChartKind, direction: 'horizontal' | 'vertical'): boolean =>
  kind === 'scatter' || (kind === 'bar') === (direction === 'horizontal');

/** Horizontal gridlines run off the vertical axis, and vertical ones off the horizontal. */
const gridlineAxis = (direction: 'horizontal' | 'vertical'): 'horizontal' | 'vertical' =>
  direction === 'horizontal' ? 'vertical' : 'horizontal';

const legendShown = (chart: DrawableChart): boolean => chart.legend?.visible !== false;

const labelsShown = (chart: DrawableChart): boolean =>
  chart.dataLabels?.visible === true ||
  (chart.seriesStyles ?? []).some((style) => style.dataLabels?.visible === true);

/** Whether `action` applies to `chart`, and whether it describes the chart as it is. */
export const chartElementState = (
  chart: SessionChart,
  action: ChartElementAction,
): { available: boolean; checked: boolean } => {
  if (!isDrawable(chart)) return { available: false, checked: false };
  const { kind } = chart;
  switch (action) {
    case 'axis-title-horizontal':
    case 'axis-title-vertical': {
      const axis = chart[axisKey(action === 'axis-title-horizontal' ? 'horizontal' : 'vertical')];
      return { available: hasAxes(kind), checked: hasAxes(kind) && axis?.title !== undefined };
    }
    case 'gridlines-horizontal':
    case 'gridlines-vertical': {
      const from = gridlineAxis(action === 'gridlines-horizontal' ? 'horizontal' : 'vertical');
      const available = hasAxes(kind) && isValueAxis(kind, from);
      return { available, checked: available && chart[axisKey(from)]?.majorGridlines !== false };
    }
    case 'title-none':
      return { available: true, checked: chart.title === '' };
    case 'title-above':
      return { available: true, checked: chart.title !== '' };
    case 'labels-none':
      return { available: true, checked: !labelsShown(chart) };
    case 'labels-show':
      return { available: true, checked: labelsShown(chart) };
    case 'legend-none':
      return { available: true, checked: !legendShown(chart) };
    default: {
      const position = action.slice('legend-'.length);
      return {
        available: true,
        checked: legendShown(chart) && (chart.legend?.position ?? 'bottom') === position,
      };
    }
  }
};

const toggledAxis = (
  axis: ChartAxisSpec | undefined,
  patch: Partial<ChartAxisSpec>,
): ChartAxisSpec => ({ ...axis, ...patch });

/** The change `action` makes to `chart`, or null when it does not apply. */
export const chartElementPatch = (
  chart: SessionChart,
  action: ChartElementAction,
): SessionChartPatch | null => {
  const state = chartElementState(chart, action);
  if (!state.available || !isDrawable(chart)) return null;
  switch (action) {
    case 'axis-title-horizontal':
    case 'axis-title-vertical': {
      const key = axisKey(action === 'axis-title-horizontal' ? 'horizontal' : 'vertical');
      const { title: _title, ...rest } = chart[key] ?? {};
      return { [key]: state.checked ? rest : toggledAxis(rest, { title: { auto: true } }) };
    }
    case 'gridlines-horizontal':
    case 'gridlines-vertical': {
      const key = axisKey(
        gridlineAxis(action === 'gridlines-horizontal' ? 'horizontal' : 'vertical'),
      );
      return { [key]: toggledAxis(chart[key], { majorGridlines: !state.checked }) };
    }
    case 'title-none':
      return state.checked ? null : { title: '' };
    case 'title-above':
      return state.checked ? null : { title: { auto: true } };
    case 'labels-none':
    case 'labels-show': {
      if (state.checked) return null;
      const visible = action === 'labels-show';
      return {
        dataLabels: {
          ...chart.dataLabels,
          visible,
          ...(visible ? { showValue: chart.dataLabels?.showValue ?? true } : {}),
        },
        // A series showing labels of its own follows the chart.
        seriesStyles: chart.seriesStyles?.map((style) =>
          style.dataLabels ? { ...style, dataLabels: { ...style.dataLabels, visible } } : style,
        ),
      };
    }
    case 'legend-none':
      return state.checked ? null : { legend: { ...chart.legend, visible: false } };
    default: {
      if (state.checked) return null;
      const position = action.slice('legend-'.length) as 'right' | 'top' | 'left' | 'bottom';
      return { legend: { visible: true, position } };
    }
  }
};

/** Add, move or remove a chart element, as the Add Chart Element menu does;
 *  false when it changes nothing or the change is refused. */
export const applySessionChartElement = (
  store: SpreadsheetStore,
  id: string,
  action: ChartElementAction,
  history: History | null,
  wb: WorkbookHandle,
): boolean => {
  const chart = sessionChartById(store.getState(), id);
  const patch = chart ? chartElementPatch(chart, action) : null;
  return patch !== null && updateSessionChart(store, id, patch, history, wb) !== null;
};

/** Plot the chart's rows as series where it plotted columns, and the reverse. */
export const switchSessionChartRowColumn = (
  store: SpreadsheetStore,
  id: string,
  history: History | null,
  wb: WorkbookHandle,
): boolean => {
  const chart = sessionChartById(store.getState(), id);
  if (!chart || !isDrawable(chart) || chart.kind === 'scatter') return false;
  const seriesDirection = chart.seriesDirection === 'rows' ? 'columns' : 'rows';
  return updateSessionChart(store, id, { seriesDirection }, history, wb) !== null;
};

/** The colours of `preset` for the chart's sheet's workbook. */
const presetColors = (preset: ChartPalettePreset, wb: WorkbookHandle): string[] =>
  chartPalettePreset(preset, accentsFromTheme(wb.getTheme()?.colors));

/** Whether the chart's series take `preset`'s colours; a series with no colour
 *  of its own takes the theme's. */
export const chartUsesPalette = (
  chart: SessionChart,
  preset: ChartPalettePreset,
  wb: WorkbookHandle,
): boolean => {
  if (!isDrawable(chart)) return false;
  const colors = presetColors(preset, wb);
  const automatic = presetColors('office', wb);
  const series = chart.data?.series ?? [];
  return (
    series.length > 0 &&
    series.every((entry, index) => {
      const style = chart.seriesStyles?.find((candidate) => candidate.seriesId === entry.id);
      const color = seriesColorOf(chart, style) ?? automatic[index % automatic.length];
      return color?.toLowerCase() === colors[index % colors.length]?.toLowerCase();
    })
  );
};

/** Colour the chart's series with `preset`, as Change Colors does. */
export const applySessionChartColors = (
  store: SpreadsheetStore,
  id: string,
  preset: ChartPalettePreset,
  history: History | null,
  wb: WorkbookHandle,
): boolean => {
  const chart = sessionChartById(store.getState(), id);
  if (!chart || !isDrawable(chart)) return false;
  const { kind } = chart;
  return (
    updateSessionChart(
      store,
      id,
      {
        palette: presetColors(preset, wb),
        // Series coloured by hand take the new palette too.
        seriesStyles: chart.seriesStyles?.map((style) => withoutSeriesColor(kind, style)),
      },
      history,
      wb,
    ) !== null
  );
};
