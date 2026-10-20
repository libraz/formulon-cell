import { chartDataFromEngine, chartPreviewOverrides } from '../charts/engine-adapter.js';
import { formatChartValue } from '../charts/format.js';
import { chartTitleText, layoutChart, layoutUnsupportedChart } from '../charts/layout.js';
import { accentsFromTheme } from '../charts/series-colors.js';
import type { ChartData, ChartKind, ChartPrimitive, ChartScene } from '../charts/types.js';
import type { History } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import {
  addSessionChartSeries,
  canMutateChart,
  clearSessionChart,
  removeSessionChartSeries,
  type SessionChartSeriesInput,
  sessionChartById,
  updateSessionChart,
} from '../commands/session-chart.js';
import {
  ANCHOR_KIND_ABSOLUTE,
  drawingClipPath,
  placementForPixels,
  sessionPosition,
} from '../engine/drawing-geometry.js';
import { rangeForRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import type { ChartSourceHighlight } from '../store/types.js';
import { createExcelRibbonSvg } from '../toolbar/excel-ribbon-icons.js';
import { createInteractionButton } from './chip-button.js';
import { formatSessionChartValue } from './session-chart-format.js';
import {
  type ChartValueFormatter,
  chartSettingsDefaults,
  createSessionChartSettings,
  type SessionChartSettingsHandle,
  type SessionChartSettingsLabels,
  type SessionChartSettingsShape,
} from './session-chart-settings.js';

export interface SessionChartsHandle {
  refresh(): void;
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
/** Pointer travel before a press on a chart becomes a move. */
const DRAG_THRESHOLD = 3;
/** Width and height the stylesheet takes from a chart frame for its border and
 *  plot padding; the scene is laid out in what remains. */
const PLOT_INSET = 14;

/** The sizing handles of a selected chart, by compass point. */
const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
type Handle = (typeof HANDLES)[number];

/** Data marks that carry a tooltip of their point. */
const POINT_ROLES = new Set(['bar', 'point', 'slice']);

const DEFAULT_LABELS: SessionChartLabels = {
  elementsButton: 'Chart Elements',
  stylesButton: 'Chart Styles',
  filtersButton: 'Chart Filters',
  pointTooltip: 'Series "{series}" Point "{point}"',
  pointValue: 'Value: {value}',
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

/** Elements that keep a press to themselves instead of starting a move. */
const INTERACTIVE = 'button, input, select, textarea, option, [contenteditable="true"]';

/** The frame `start` takes when `handle` is dragged by (dx, dy); a corner
 *  dragged with `keepRatio` scales the frame without changing its shape. */
const resizedFrame = (
  start: { x: number; y: number; w: number; h: number },
  handle: Handle,
  dx: number,
  dy: number,
  keepRatio: boolean,
): { x: number; y: number; w: number; h: number } => {
  const west = handle.endsWith('w');
  const east = handle.endsWith('e');
  const north = handle.startsWith('n');
  const south = handle.startsWith('s');
  let w = Math.max(MIN_W, east ? start.w + dx : west ? start.w - dx : start.w);
  let h = Math.max(MIN_H, south ? start.h + dy : north ? start.h - dy : start.h);
  if (keepRatio && (east || west) && (north || south)) {
    const scale = Math.max(w / start.w, h / start.h);
    w = Math.max(MIN_W, start.w * scale);
    h = Math.max(MIN_H, start.h * scale);
  }
  return {
    x: west ? start.x + start.w - w : start.x,
    y: north ? start.y + start.h - h : start.y,
    w,
    h,
  };
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

const sameHighlights = (
  a: readonly ChartSourceHighlight[],
  b: readonly ChartSourceHighlight[],
): boolean =>
  a.length === b.length &&
  a.every(
    (ref, i) =>
      ref.sheet === b[i]?.sheet &&
      ref.r0 === b[i]?.r0 &&
      ref.c0 === b[i]?.c0 &&
      ref.r1 === b[i]?.r1 &&
      ref.c1 === b[i]?.c1 &&
      ref.role === b[i]?.role,
  );

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
  host.appendChild(root);
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

  let publishingHighlight = false;
  /** Outline the selected chart's data on the grid, as Excel does. */
  const syncSourceHighlight = (): void => {
    const chart = selectedId === null ? null : sessionChartById(store.getState(), selectedId);
    const next = chart ? chartSourceHighlights(chart) : [];
    if (sameHighlights(store.getState().ui.chartSourceRefs, next)) return;
    // The outline is grid state only; the charts themselves do not change.
    publishingHighlight = true;
    try {
      mutators.setChartSourceRefs(store, next);
    } finally {
      publishingHighlight = false;
    }
  };

  const updateSelectionClasses = (): void => {
    syncSourceHighlight();
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

  /** Whether `target` is part of a chart or its settings, which keep the chart selected. */
  const insideCharts = (target: EventTarget | null): boolean =>
    target instanceof Node &&
    (root.contains(target) ||
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
    handle?: Handle,
  ): void => {
    if (e.button !== 0 || !canPlace(chart)) return;
    if (handle) e.preventDefault();
    const startX = e.clientX;
    const startY = e.clientY;
    const start = sessionPosition(store, chart.nativeAnchor);
    let latest: { x?: number; y?: number; w?: number; h?: number } | null = null;
    const preview = (patch: { x?: number; y?: number; w?: number; h?: number }): void => {
      if (patch.x !== undefined) panel.style.left = `${patch.x}px`;
      if (patch.y !== undefined) panel.style.top = `${patch.y}px`;
      if (patch.w !== undefined) panel.style.width = `${patch.w}px`;
      if (patch.h !== undefined) panel.style.height = `${patch.h}px`;
      latest = patch;
    };
    const onMove = (move: PointerEvent): void => {
      const dx = move.clientX - startX;
      const dy = move.clientY - startY;
      if (!latest && Math.hypot(dx, dy) < DRAG_THRESHOLD) return;
      move.preventDefault();
      if (!handle) {
        preview({ x: start.x + dx, y: start.y + dy });
        return;
      }
      const frame = resizedFrame(start, handle, dx, dy, move.shiftKey);
      preview({
        ...(frame.x === start.x ? {} : { x: frame.x }),
        ...(frame.y === start.y ? {} : { y: frame.y }),
        w: frame.w,
        h: frame.h,
      });
    };
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      if (latest && !placeChart(chart, latest)) render();
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp, { once: true });
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
    root.replaceChildren();
    // The store keeps a sheet's charts in the workbook's stacking order.
    const charts = state.charts.charts.filter((chart) => chart.sheet === state.data.sheetIndex);
    if (selectedId !== null && !charts.some((chart) => chart.id === selectedId)) selectedId = null;
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
          ? 'ArrowLeft ArrowRight ArrowUp ArrowDown Shift+ArrowLeft Shift+ArrowRight Shift+ArrowUp Shift+ArrowDown Delete Backspace Escape'
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
        for (const handle of HANDLES) {
          const grip = document.createElement('div');
          grip.className = 'fc-chart__handle';
          grip.dataset.handle = handle;
          grip.setAttribute('aria-hidden', 'true');
          grip.addEventListener('pointerdown', (e) => {
            e.stopPropagation();
            selectChart(chart.id, panel);
            applyDrag(e, chart, panel, handle);
          });
          panel.appendChild(grip);
        }
      }
      root.appendChild(panel);
      if (drawable) settings.sync(chart, chart.data, panel);
    });
    if (settings.isOpen() && !visibleChartIds.has(settings.chartId() ?? '')) {
      settings.close(true);
    }
    // The selected chart may have been removed or given another data range.
    syncSourceHighlight();
  };

  const unsubscribe = store.subscribe(() => {
    if (!publishingHighlight) render();
  });
  render();

  return {
    refresh: render,
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
      if (store.getState().ui.chartSourceRefs.length > 0) mutators.setChartSourceRefs(store, []);
      unsubscribeController();
      settings.detach();
      root.remove();
    },
  };
}
