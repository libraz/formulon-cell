import { chartDataFromEngine, chartPreviewOverrides } from '../charts/engine-adapter.js';
import { formatChartValue } from '../charts/format.js';
import { chartTitleText, layoutChart, layoutUnsupportedChart } from '../charts/layout.js';
import { accentsFromTheme } from '../charts/series-colors.js';
import type { ChartData, ChartKind, ChartPrimitive, ChartScene } from '../charts/types.js';
import type { History } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import {
  addSessionChartSeries,
  arrangeSessionChart,
  canMutateChart,
  clearSessionChart,
  duplicateSessionChart,
  removeSessionChartSeries,
  type SessionChartSeriesInput,
  sessionChartById,
  updateSessionChart,
} from '../commands/session-chart.js';
import {
  ANCHOR_KIND_ABSOLUTE,
  drawingClipPath,
  nearestGridline,
  placementForPixels,
  sessionPosition,
} from '../engine/drawing-geometry.js';
import { rangeForRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import type { ChartSourceHighlight, State } from '../store/types.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { createInteractionButton } from './chip-button.js';
import {
  appendDrawingHandles,
  createObjectSelectionPublisher,
  type DrawingHandle,
  focusedPanelId,
  focusPanel,
  INTERACTIVE,
  isDuplicateShortcut,
  keepKeyFromGrid,
  keepsObjectSelected,
  previewFrame,
  trackFrameDrag,
} from './drawing-frame.js';
import { createObjectMenu, type ObjectMenuEntry } from './object-menu.js';
import { formatSessionChartValue } from './session-chart-format.js';
import {
  type ChartValueFormatter,
  chartSettingsDefaults,
  createSessionChartSettings,
  type SessionChartSettingsHandle,
  type SessionChartSettingsLabels,
  type SessionChartSettingsShape,
  type SettingsTab,
} from './session-chart-settings.js';

export interface SessionChartsHandle {
  refresh(): void;
  /** Open the settings of the selected chart on `tab`; false when no chart
   *  that can be drawn is selected. */
  openSettings(tab: SettingsTab): boolean;
  setLabels(labels: Partial<SessionChartLabels>, locale?: string): void;
  detach(): void;
}

export interface SessionChartLabels extends Partial<SessionChartSettingsLabels> {
  /** The buttons beside a selected chart. */
  elementsButton: string;
  stylesButton: string;
  filtersButton: string;
  /** Data point tooltip; `{series}` and `{point}` are the series and category names. */
  pointTooltip: string;
  /** Second tooltip line; `{value}` is the formatted value. */
  pointValue: string;
  /** The name box text for a selected chart; `{n}` is its number on the sheet. */
  objectName: string;
  /** The chart's right-click menu. */
  changeChartType: string;
  selectData: string;
  formatChartArea: string;
  bringToFront: string;
  sendToBack: string;
  columnChart: string;
  barChart: string;
  lineChart: string;
  areaChart: string;
  pieChart: string;
  doughnutChart: string;
  scatterChart: string;
  unsupportedChart: string;
  /** The automatic title of a chart with several series. */
  chartTitle: string;
  /** The automatic title of an axis. */
  axisTitle: string;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const MIN_W = 240;
const MIN_H = 160;
const KEYBOARD_STEP = 8;
/** Width and height the stylesheet takes from a chart frame for its border and
 *  plot padding; the scene is laid out in what remains. */
const PLOT_INSET = 14;

/** Data marks that carry a tooltip of their point. */
const POINT_ROLES = new Set(['bar', 'point', 'slice']);

const DEFAULT_LABELS: SessionChartLabels = {
  elementsButton: 'Chart Elements',
  stylesButton: 'Chart Styles',
  filtersButton: 'Chart Filters',
  pointTooltip: 'Series "{series}" Point "{point}"',
  pointValue: 'Value: {value}',
  objectName: 'Chart {n}',
  changeChartType: 'Change Chart Type...',
  selectData: 'Select Data...',
  formatChartArea: 'Format Chart Area...',
  bringToFront: 'Bring to Front',
  sendToBack: 'Send to Back',
  columnChart: 'Column chart',
  barChart: 'Bar chart',
  lineChart: 'Line chart',
  areaChart: 'Area chart',
  pieChart: 'Pie chart',
  doughnutChart: 'Doughnut chart',
  scatterChart: 'Scatter chart',
  unsupportedChart: 'This chart type cannot be displayed.',
  chartTitle: 'Chart Title',
  axisTitle: 'Axis Title',
};

const SETTINGS_TAB_LABELS: Pick<
  SessionChartSettingsLabels,
  'elementsTab' | 'stylesTab' | 'filtersTab'
> = {
  elementsTab: chartSettingsDefaults.elementsTab,
  stylesTab: chartSettingsDefaults.stylesTab,
  filtersTab: chartSettingsDefaults.filtersTab,
};

const isDrawable = (chart: SessionChart): chart is SessionChartSettingsShape =>
  chart.kind !== 'unsupported';

const chartLabel = (chart: SessionChart, labels: SessionChartLabels): string => {
  if (chart.kind === 'unsupported') return chart.unsupported?.title || labels.unsupportedChart;
  const title = chartTitleText(chart.title, chart.data, labels.chartTitle);
  if (title) return title;
  if (chart.kind === 'bar') return labels.barChart;
  if (chart.kind === 'line') return labels.lineChart;
  if (chart.kind === 'area') return labels.areaChart;
  if (chart.kind === 'pie') return labels.pieChart;
  if (chart.kind === 'doughnut') return labels.doughnutChart;
  if (chart.kind === 'scatter') return labels.scatterChart;
  return labels.columnChart;
};

/** The series as a draft of another chart family would plot them: Excel uses
 *  numeric categories as X values (else 1, 2, 3, ...) and X values as categories. */
const dataForKind = (data: ChartData, kind: ChartKind): ChartData => {
  if ((kind === 'scatter') === (data.domain === 'xy')) return data;
  if (data.domain === 'category') {
    const numeric = data.categories.map((label) => (label.trim() === '' ? Number.NaN : +label));
    const useLabels = numeric.every(Number.isFinite);
    return {
      domain: 'xy',
      series: data.series.map((series) => ({
        id: series.id,
        ...(series.name === undefined ? {} : { name: series.name }),
        points: series.values.map((y, index) => ({
          id: `point-${index}`,
          x: useLabels ? (numeric[index] ?? null) : index + 1,
          y,
          ...(series.states?.[index] === undefined ? {} : { state: series.states[index] }),
        })),
      })),
    };
  }
  const first = data.series[0]?.points ?? [];
  return {
    domain: 'category',
    categories: first.map((point, index) => (point.x === null ? String(index + 1) : `${point.x}`)),
    categoryIds: first.map((_point, index) => `category-${index}`),
    series: data.series.map((series) => ({
      id: series.id,
      ...(series.name === undefined ? {} : { name: series.name }),
      values: series.points.map((point) => point.y),
      states: series.points.map((point) => point.state ?? 'value'),
    })),
  };
};

/** The parts of a chart's data range Excel outlines while it is selected:
 *  series names, categories (or X values) and values. */
export const chartSourceHighlights = (chart: SessionChart): ChartSourceHighlight[] => {
  const source = chart.source;
  if (!source || chart.kind === 'unsupported') return [];
  const headerRow = chart.headerRow === true;
  const headerColumn = chart.headerColumn === true;
  const top = source.r0 + (headerRow ? 1 : 0);
  const left = source.c0 + (headerColumn ? 1 : 0);
  const part = (
    role: ChartSourceHighlight['role'],
    r0: number,
    c0: number,
    r1: number,
    c1: number,
  ): ChartSourceHighlight[] =>
    r0 <= r1 && c0 <= c1 ? [{ sheet: source.sheet, r0, c0, r1, c1, role }] : [];
  const byRows = chart.seriesDirection === 'rows';
  const names = byRows
    ? headerColumn
      ? part('name', top, source.c0, source.r1, source.c0)
      : []
    : headerRow
      ? part('name', source.r0, left, source.r0, source.c1)
      : [];
  const categories = byRows
    ? headerRow
      ? part('categories', source.r0, left, source.r0, source.c1)
      : []
    : headerColumn
      ? part('categories', top, source.c0, source.r1, source.c0)
      : [];
  return [...names, ...categories, ...part('values', top, left, source.r1, source.c1)];
};

/** Excel's name for a chart, numbered in the order of its sheet's charts. */
const chartObjectName = (state: State, chart: SessionChart, labels: SessionChartLabels): string => {
  const index = state.charts.charts
    .filter((entry) => entry.sheet === chart.sheet)
    .findIndex((entry) => entry.id === chart.id);
  return labels.objectName.replace('{n}', String(index + 1));
};

const appendChartPrimitive = (parent: SVGSVGElement, primitive: ChartPrimitive): void => {
  const element = document.createElementNS(SVG_NS, primitive.tag);
  for (const [key, value] of Object.entries(primitive.attrs)) {
    element.setAttribute(key, String(value));
  }
  if (primitive.role !== undefined) element.setAttribute('data-chart-role', primitive.role);
  if (primitive.seriesIndex !== undefined)
    element.setAttribute('data-series-index', String(primitive.seriesIndex));
  if (primitive.pointIndex !== undefined)
    element.setAttribute('data-point-index', String(primitive.pointIndex));
  if (primitive.seriesId !== undefined) element.setAttribute('data-series-id', primitive.seriesId);
  if (primitive.pointId !== undefined) element.setAttribute('data-point-id', primitive.pointId);
  if (primitive.text !== undefined) element.textContent = primitive.text;
  parent.appendChild(element);
};

/** Map the pure chart scene to native SVG nodes. Text is assigned through
 * `textContent`; chart data never becomes markup or an HTML string. */
const renderChartScene = (svg: SVGSVGElement, scene: ChartScene): void => {
  svg.replaceChildren();
  svg.setAttribute('viewBox', `0 0 ${scene.width} ${scene.height}`);
  for (const primitive of scene.elements) appendChartPrimitive(svg, primitive);
};

export function attachSessionCharts(deps: {
  host: HTMLElement;
  /** The element the charts are drawn over, whose top-left corner is the
   *  grid's; defaults to `host`. */
  layer?: HTMLElement;
  store: SpreadsheetStore;
  /** The workbook the charts live in; without one the overlays are read-only. */
  getWb?: () => WorkbookHandle | null;
  labels?: Partial<SessionChartLabels>;
  history?: History | null;
  locale?: string;
  formatValue?: ChartValueFormatter;
}): SessionChartsHandle {
  const { host, store, history = null, getWb = (): WorkbookHandle | null => null } = deps;
  const formatValue = deps.formatValue ?? formatSessionChartValue;
  let locale = deps.locale;
  let labels: SessionChartLabels = { ...DEFAULT_LABELS, ...deps.labels };
  const root = document.createElement('div');
  root.className = 'fc-charts';
  (deps.layer ?? host).appendChild(root);
  let selectedId: string | null = null;
  let settings: SessionChartSettingsHandle;

  /** The workbook when it can change charts, else null. */
  const chartWorkbook = (): WorkbookHandle | null => {
    const wb = getWb();
    return wb?.capabilities.charts ? wb : null;
  };

  /** The workbook theme's accents, which colour series without a colour of their own. */
  const workbookAccents = (): readonly string[] | undefined =>
    accentsFromTheme(getWb()?.getTheme()?.colors);
  const themeAccents = (): { accents?: readonly string[] } => {
    const accents = workbookAccents();
    return accents ? { accents } : {};
  };

  const canEditChart = (chart: SessionChart): boolean => {
    if (!chartWorkbook()) return false;
    const { fromRow, fromCol } = chart.nativeAnchor;
    const cells = [{ sheet: chart.sheet, row: fromRow, col: fromCol }];
    if (chart.source) {
      cells.push({ sheet: chart.source.sheet, row: chart.source.r0, col: chart.source.c0 });
    }
    return canMutateChart(store, {
      operation: 'object',
      origin: 'instanceApi',
      effects: cells.map((cell) => ({ kind: 'cells', cells: [cell] })),
    });
  };

  /** Series for a settings draft, resolved by the engine as the draft would split them. */
  const draftData = (chart: SessionChartSettingsShape): ChartData => {
    const preview = getWb()?.previewChartData(
      chart.sheet,
      chart.nativeObjectId,
      chartPreviewOverrides(chart),
    );
    return dataForKind(preview ? chartDataFromEngine(preview) : chart.data, chart.kind);
  };

  settings = createSessionChartSettings({
    host,
    store,
    history,
    locale,
    formatValue,
    labels: { ...SETTINGS_TAB_LABELS, ...deps.labels },
    canEdit: canEditChart,
    readData: draftData,
    accents: workbookAccents,
    onApply: (id, patch, patchHistory) => {
      const wb = chartWorkbook();
      if (wb) updateSessionChart(store, id, patch, patchHistory, wb);
    },
    onAddSeries: (id, refs, seriesHistory) => {
      const wb = chartWorkbook();
      const chart = sessionChartById(store.getState(), id);
      if (!wb || !chart) return null;
      // A reference without a sheet is on the chart's sheet.
      const resolve = (ref: string | undefined): Range | null | undefined =>
        ref === undefined ? undefined : rangeForRef(wb, ref, chart.sheet);
      const values = resolve(refs.values);
      const name = resolve(refs.name);
      const categories = resolve(refs.categories);
      const x = resolve(refs.x);
      if (!values || name === null || categories === null || x === null) return null;
      const series: SessionChartSeriesInput = {
        values,
        ...(name ? { name } : {}),
        ...(categories ? { categories } : {}),
        ...(x ? { x } : {}),
      };
      const before = new Set(chart.data.series.map((entry) => entry.id));
      const updated = addSessionChartSeries(store, id, series, seriesHistory, wb);
      return updated?.data.series.find((entry) => !before.has(entry.id))?.id ?? null;
    },
    onRemoveSeries: (id, seriesId, seriesHistory) => {
      const wb = chartWorkbook();
      return (
        wb !== null && removeSessionChartSeries(store, id, seriesId, seriesHistory, wb) !== null
      );
    },
  });

  const selection = createObjectSelectionPublisher(store);
  /** Publish the selected chart to the name box and outline its data on the
   *  grid, as Excel does. */
  const syncSelection = (): void => {
    const state = store.getState();
    const chart = selectedId === null ? null : sessionChartById(state, selectedId);
    selection.publish(
      chart ? { id: chart.id, name: chartObjectName(state, chart, labels) } : null,
      chart ? chartSourceHighlights(chart) : [],
    );
  };

  const updateSelectionClasses = (): void => {
    syncSelection();
    const panels = Array.from(root.querySelectorAll<HTMLElement>('.fc-chart'));
    for (const [index, panel] of panels.entries()) {
      const selected = panel.dataset.chartId === selectedId;
      panel.classList.toggle('fc-chart--selected', selected);
      panel.setAttribute('aria-selected', selected ? 'true' : 'false');
      panel.style.zIndex = selected ? String(panels.length + 1) : String(index + 1);
    }
  };

  const selectChart = (chartId: string, panel: HTMLElement): void => {
    selectedId = chartId;
    updateSelectionClasses();
    if (document.activeElement !== panel) panel.focus({ preventScroll: true });
  };

  const deselect = (): void => {
    if (selectedId === null) return;
    selectedId = null;
    updateSelectionClasses();
  };

  /** Take up a chart another part of the sheet selected, and leave the chart
   *  when another object is selected; returns the id taken up. */
  const adoptSelection = (state: State, charts: readonly SessionChart[]): string | null => {
    const wanted = state.ui.selectedObject?.id;
    if (wanted === undefined || wanted === selectedId) return null;
    if (!charts.some((chart) => chart.id === wanted)) {
      selectedId = null;
      return null;
    }
    selectedId = wanted;
    return wanted;
  };

  /** Whether `target` is part of a chart, its settings or the ribbon, which keep the chart selected. */
  const menu = createObjectMenu(host);
  const insideCharts = (target: EventTarget | null): boolean =>
    target instanceof Node &&
    (root.contains(target) ||
      menu.contains(target) ||
      keepsObjectSelected(target) ||
      (target instanceof Element && target.closest('.fc-chart-settings') !== null));

  // Like Excel, a press anywhere else on the sheet leaves the chart.
  const onDocumentPointerDown = (e: PointerEvent): void => {
    if (!insideCharts(e.target)) deselect();
  };
  document.addEventListener('pointerdown', onDocumentPointerDown, true);
  const onFocusOut = (e: FocusEvent): void => {
    if (e.relatedTarget !== null && !insideCharts(e.relatedTarget)) deselect();
  };
  root.addEventListener('focusout', onFocusOut);
  root.addEventListener('keydown', keepKeyFromGrid);

  /** Move or resize a chart to grid pixels; false when the workbook refused. */
  const placeChart = (
    chart: SessionChart,
    patch: { x?: number; y?: number; w?: number; h?: number },
  ): boolean => {
    const wb = chartWorkbook();
    if (!wb) return false;
    const anchor = placementForPixels(store, chart.nativeAnchor, patch);
    return updateSessionChart(store, chart.id, { anchor }, history, wb) !== null;
  };

  const removeChart = (chart: SessionChart): void => {
    const wb = chartWorkbook();
    if (wb) clearSessionChart(store, chart.id, history, wb);
    host.focus({ preventScroll: true });
  };

  /** The engine cannot re-anchor an absolutely positioned object. */
  const canPlace = (chart: SessionChart): boolean =>
    chart.nativeAnchor.anchorKind !== ANCHOR_KIND_ABSOLUTE;

  /** Drag the whole chart (`handle` omitted) or one of its sizing handles. The
   *  frame follows the pointer and the workbook is changed once, on release. */
  const applyDrag = (
    e: PointerEvent,
    chart: SessionChart,
    panel: HTMLElement,
    handle?: DrawingHandle,
  ): void => {
    if (e.button !== 0 || !canPlace(chart)) return;
    if (handle) e.preventDefault();
    trackFrameDrag(e, {
      start: sessionPosition(store, chart.nativeAnchor),
      ...(handle ? { handle } : {}),
      min: { w: MIN_W, h: MIN_H },
      snap: (axis, value) => nearestGridline(store, axis, value),
      keepRatio: (move) => move.shiftKey,
      preview: (patch) => previewFrame(panel, patch),
      commit: (patch) => {
        if (!placeChart(chart, patch)) render();
      },
    });
  };

  /** Open the settings surface on `tab`, beside the chart's buttons. */
  const openSettings = (chart: SessionChart, panel: HTMLElement, tab: SettingsTab): void => {
    const anchor = panel.querySelector<HTMLElement>(`[data-chart-settings-tab="${tab}"]`);
    if (anchor && isDrawable(chart)) settings.open(chart, chart.data, anchor, tab);
  };

  /** The chart's right-click menu, in the spreadsheet's order. */
  const openMenu = (chart: SessionChart, panel: HTMLElement, x: number, y: number): void => {
    const state = store.getState();
    const sheetCharts = state.charts.charts.filter((entry) => entry.sheet === chart.sheet);
    const rank = sheetCharts.findIndex((entry) => entry.id === chart.id);
    const arrangeable = chartWorkbook() !== null && canEditChart(chart);
    const arrange = (action: 'bring-front' | 'send-back'): void => {
      const wb = chartWorkbook();
      if (wb) arrangeSessionChart(store, chart.id, action, history, wb);
    };
    const entries: ObjectMenuEntry[] = [];
    if (isDrawable(chart)) {
      entries.push(
        {
          id: 'changeChartType',
          label: labels.changeChartType,
          run: () => openSettings(chart, panel, 'elements'),
        },
        {
          id: 'selectData',
          label: labels.selectData,
          run: () => openSettings(chart, panel, 'filters'),
        },
        { separator: true },
      );
    }
    entries.push(
      {
        id: 'bringToFront',
        label: labels.bringToFront,
        disabled: !arrangeable || rank === sheetCharts.length - 1,
        run: () => arrange('bring-front'),
      },
      {
        id: 'sendToBack',
        label: labels.sendToBack,
        disabled: !arrangeable || rank === 0,
        run: () => arrange('send-back'),
      },
    );
    if (isDrawable(chart)) {
      entries.push(
        { separator: true },
        {
          id: 'formatChartArea',
          label: labels.formatChartArea,
          run: () => openSettings(chart, panel, 'styles'),
        },
      );
    }
    menu.open(x, y, entries, {
      label: chartObjectName(state, chart, labels),
      restoreFocusTo: panel,
    });
  };

  const applyKeyboard = (e: KeyboardEvent, chart: SessionChart, panel: HTMLElement): void => {
    if (e.defaultPrevented) return;
    const target = e.target;
    if (target instanceof HTMLElement && target.closest(INTERACTIVE)) return;
    if (e.key === 'Escape') {
      // An open settings surface takes Escape for itself.
      if (settings.isOpen()) return;
      e.preventDefault();
      deselect();
      host.focus({ preventScroll: true });
      return;
    }
    selectChart(chart.id, panel);
    if (e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey)) {
      e.preventDefault();
      const rect = panel.getBoundingClientRect();
      openMenu(chart, panel, rect.left + 8, rect.top + 8);
      return;
    }
    if (isDuplicateShortcut(e)) {
      e.preventDefault();
      const wb = chartWorkbook();
      const copy = wb ? duplicateSessionChart(store, chart.id, history, wb) : null;
      // Like the spreadsheet, the copy takes the selection.
      const copyPanel = Array.from(root.querySelectorAll<HTMLElement>('.fc-chart')).find(
        (candidate) => candidate.dataset.chartId === copy?.id,
      );
      if (copy && copyPanel) selectChart(copy.id, copyPanel);
      return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      removeChart(chart);
      return;
    }

    const deltaByKey: Record<string, [number, number]> = {
      ArrowLeft: [-KEYBOARD_STEP, 0],
      ArrowRight: [KEYBOARD_STEP, 0],
      ArrowUp: [0, -KEYBOARD_STEP],
      ArrowDown: [0, KEYBOARD_STEP],
    };
    const delta = deltaByKey[e.key];
    if (!delta) return;
    e.preventDefault();
    if (!canPlace(chart)) return;

    const [dx, dy] = delta;
    const current = sessionPosition(store, chart.nativeAnchor);
    if (e.shiftKey) {
      placeChart(chart, {
        w: Math.max(MIN_W, current.w + dx),
        h: Math.max(MIN_H, current.h + dy),
      });
      return;
    }
    placeChart(chart, { x: current.x + dx, y: current.y + dy });
  };

  /** The tooltip of one data point, as Excel words it. */
  const pointTooltip = (chart: SessionChart, seriesIndex: number, pointIndex: number): string => {
    const { data } = chart;
    const series = data.series[seriesIndex];
    if (!series) return '';
    const seriesName = series.name ?? `Series${seriesIndex + 1}`;
    const format = (value: number | null): string =>
      value === null
        ? ''
        : (formatChartValue(value, series.formatCode, locale ?? 'en-US', formatValue) ??
          String(value));
    let point: string;
    let value: string;
    if (data.domain === 'category') {
      const raw = data.series[seriesIndex]?.values[pointIndex] ?? null;
      point = data.categories[pointIndex] ?? String(pointIndex + 1);
      value = format(raw);
      if ((chart.kind === 'pie' || chart.kind === 'doughnut') && raw !== null) {
        const values = data.series[seriesIndex]?.values ?? [];
        const total = values.reduce<number>((sum, entry) => sum + Math.abs(entry ?? 0), 0);
        if (total > 0) value += ` (${Math.round((Math.abs(raw) / total) * 100)}%)`;
      }
    } else {
      const entry = data.series[seriesIndex]?.points[pointIndex];
      if (!entry) return '';
      point = entry.x === null ? String(pointIndex + 1) : format(entry.x);
      value = `(${point}, ${format(entry.y)})`;
    }
    const heading = labels.pointTooltip.replace('{series}', seriesName).replace('{point}', point);
    return `${heading}\n${labels.pointValue.replace('{value}', value)}`;
  };

  const addPointTooltips = (svg: SVGSVGElement, chart: SessionChart): void => {
    for (const mark of Array.from(svg.querySelectorAll<SVGElement>('[data-point-index]'))) {
      if (!POINT_ROLES.has(mark.getAttribute('data-chart-role') ?? '')) continue;
      const text = pointTooltip(
        chart,
        Number(mark.getAttribute('data-series-index')),
        Number(mark.getAttribute('data-point-index')),
      );
      if (!text) continue;
      const tooltip = document.createElementNS(SVG_NS, 'title');
      tooltip.textContent = text;
      mark.appendChild(tooltip);
    }
  };

  let subscribedController: ReturnType<typeof interactionControllerFor>;
  let unsubscribeController: () => void = () => {};
  const syncControllerSubscription = (): void => {
    const nextController = interactionControllerFor(store);
    if (nextController === subscribedController) return;
    unsubscribeController();
    subscribedController = nextController;
    unsubscribeController = nextController?.subscribe(() => render()) ?? (() => {});
  };

  const render = (): void => {
    // Policy changes are emitted by the interaction controller rather than
    // the workbook store. Keep the settings surface live, and follow a
    // controller replacement the next time the chart root renders.
    syncControllerSubscription();
    const state = store.getState();
    const visibleChartIds = new Set<string>();
    const focusedId = focusedPanelId(root, '.fc-chart', 'chartId');
    root.replaceChildren();
    // The store keeps a sheet's charts in the workbook's stacking order.
    const charts = state.charts.charts.filter((chart) => chart.sheet === state.data.sheetIndex);
    if (selectedId !== null && !charts.some((chart) => chart.id === selectedId)) selectedId = null;
    // A paste selects the chart it adds.
    const held = selectedId;
    const adopted = adoptSelection(state, charts);
    // Another object took the selection, and the focus with it.
    const left = held !== null && selectedId === null;
    charts.forEach((chart, index) => {
      const drawable = isDrawable(chart);
      if (drawable) visibleChartIds.add(chart.id);
      const position = sessionPosition(store, chart.nativeAnchor);
      const placeable = canPlace(chart);
      const panel = document.createElement('section');
      panel.className = drawable ? 'fc-chart' : 'fc-chart fc-chart--unsupported';
      panel.dataset.chartId = chart.id;
      panel.tabIndex = 0;
      panel.setAttribute('role', 'group');
      panel.setAttribute('aria-roledescription', 'chart');
      panel.setAttribute('aria-selected', chart.id === selectedId ? 'true' : 'false');
      panel.setAttribute(
        'aria-keyshortcuts',
        placeable
          ? 'ArrowLeft ArrowRight ArrowUp ArrowDown Shift+ArrowLeft Shift+ArrowRight Shift+ArrowUp Shift+ArrowDown Control+D Meta+D Delete Backspace Escape'
          : 'Delete Backspace Escape',
      );
      const label = chartLabel(chart, labels);
      panel.setAttribute('aria-label', label);
      panel.style.left = `${position.x}px`;
      panel.style.top = `${position.y}px`;
      panel.style.width = `${position.w}px`;
      panel.style.height = `${position.h}px`;
      panel.style.clipPath = drawingClipPath(store, chart.nativeAnchor, position);
      panel.style.zIndex = chart.id === selectedId ? String(charts.length + 1) : String(index + 1);
      panel.classList.toggle('fc-chart--selected', chart.id === selectedId);
      panel.classList.toggle('fc-chart--fixed', !placeable);
      panel.addEventListener('focus', () => {
        selectedId = chart.id;
        updateSelectionClasses();
      });
      panel.addEventListener('pointerdown', (e) => {
        if (e.target instanceof Element && e.target.closest(INTERACTIVE)) return;
        selectChart(chart.id, panel);
        applyDrag(e, chart, panel);
      });
      panel.addEventListener('keydown', (e) => applyKeyboard(e, chart, panel));
      // The object's own menu, not the cell menu under it.
      panel.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        e.stopPropagation();
        selectChart(chart.id, panel);
        openMenu(chart, panel, e.clientX, e.clientY);
      });

      const svg = document.createElementNS(SVG_NS, 'svg');
      svg.classList.add('fc-chart__plot');
      svg.setAttribute('role', 'img');
      svg.setAttribute('aria-label', label);
      const width = Math.max(1, position.w - PLOT_INSET);
      const height = Math.max(1, position.h - PLOT_INSET);
      panel.appendChild(svg);
      if (drawable) {
        renderChartScene(
          svg,
          layoutChart({
            spec: chart,
            data: chart.data,
            width,
            height,
            ...(locale === undefined ? {} : { locale }),
            formatValue,
            chartTitle: labels.chartTitle,
            axisTitle: labels.axisTitle,
            // Series without a colour of their own take Excel's automatic colours in the workbook theme.
            ...themeAccents(),
          }),
        );
        addPointTooltips(svg, chart);
        const buttons = document.createElement('div');
        buttons.className = 'fc-chart__buttons';
        const settingButtons: readonly {
          id: Parameters<SessionChartSettingsHandle['open']>[3];
          label: string;
          icon: string;
        }[] = [
          { id: 'elements', label: labels.elementsButton, icon: 'chartElements' },
          { id: 'styles', label: labels.stylesButton, icon: 'chartStyles' },
          { id: 'filters', label: labels.filtersButton, icon: 'chartFilters' },
        ];
        for (const item of settingButtons) {
          if (!item.id) continue;
          const button = createInteractionButton({
            className: 'fc-chart__button',
            ariaLabel: item.label,
          });
          button.title = item.label;
          const icon = createExcelRibbonSvg(item.icon, 'fc-chart__button-icon');
          if (icon) button.appendChild(icon);
          button.dataset.chartSettingsTab = item.id;
          button.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            settings.open(chart, chart.data, button, item.id);
          });
          buttons.appendChild(button);
        }
        panel.appendChild(buttons);
        // Excel opens the chart's formatting on a double click.
        panel.addEventListener('dblclick', (event) => {
          if (event.target instanceof Element && event.target.closest(INTERACTIVE)) return;
          const anchor = buttons.querySelector<HTMLButtonElement>('[data-chart-settings-tab]');
          if (anchor) settings.open(chart, chart.data, anchor, 'elements');
        });
      } else {
        renderChartScene(
          svg,
          layoutUnsupportedChart({
            ...(chart.unsupported?.title ? { title: chart.unsupported.title } : {}),
            message: labels.unsupportedChart,
            width,
            height,
          }),
        );
      }
      if (placeable) {
        appendDrawingHandles(panel, (e, handle) => {
          selectChart(chart.id, panel);
          applyDrag(e, chart, panel, handle);
        });
      }
      root.appendChild(panel);
      if (drawable) settings.sync(chart, chart.data, panel);
    });
    if (settings.isOpen() && !visibleChartIds.has(settings.chartId() ?? '')) {
      settings.close(true);
    }
    // The selected chart may have been removed or given another data range.
    syncSelection();
    // Rebuilding the panels drops focus; keep it on the chart, or on the sheet.
    const refocus = adopted ?? (left ? undefined : focusedId);
    if (refocus !== undefined) focusPanel(root, '.fc-chart', 'chartId', refocus, host);
  };

  const unsubscribe = store.subscribe(() => {
    if (!selection.publishing) render();
  });
  render();

  return {
    refresh: render,
    openSettings(tab) {
      const chart = selectedId === null ? null : sessionChartById(store.getState(), selectedId);
      const panel = Array.from(root.querySelectorAll<HTMLElement>('.fc-chart')).find(
        (candidate) => candidate.dataset.chartId === chart?.id,
      );
      if (!chart || !panel || !isDrawable(chart)) return false;
      openSettings(chart, panel, tab);
      return true;
    },
    setLabels(next, nextLocale) {
      labels = { ...labels, ...next };
      if (nextLocale) {
        // Keep the pure chart scene and the settings surface on the same
        // locale after a host switches languages.
        locale = nextLocale;
      }
      settings.setLabels(next, nextLocale);
      render();
    },
    detach() {
      unsubscribe();
      document.removeEventListener('pointerdown', onDocumentPointerDown, true);
      root.removeEventListener('focusout', onFocusOut);
      root.removeEventListener('keydown', keepKeyFromGrid);
      menu.detach();
      selectedId = null;
      syncSelection();
      unsubscribeController();
      settings.detach();
      root.remove();
    },
  };
}
