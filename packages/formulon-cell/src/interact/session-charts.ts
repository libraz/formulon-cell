import { chartDataFromEngine, chartPreviewOverrides } from '../charts/engine-adapter.js';
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
  placementForPixels,
  sessionPosition,
} from '../engine/drawing-geometry.js';
import { rangeForRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
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
  close: string;
  resize: string;
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
/** Width and height the stylesheet takes from a chart frame for its border,
 *  header and plot padding; the scene is laid out in what remains. */
const PLOT_INSET_W = 26;
const PLOT_INSET_H = 57;

const DEFAULT_LABELS: SessionChartLabels = {
  close: 'Close',
  resize: 'Resize chart',
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

const createSessionChartCloseButton = (label: string): HTMLButtonElement =>
  createInteractionButton({
    className: 'fc-chart__close',
    ariaLabel: label,
    text: '×',
  });

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
  closeLabel?: string;
  labels?: Partial<SessionChartLabels>;
  history?: History | null;
  locale?: string;
  formatValue?: ChartValueFormatter;
}): SessionChartsHandle {
  const { host, store, history = null, getWb = (): WorkbookHandle | null => null } = deps;
  const formatValue = deps.formatValue ?? formatSessionChartValue;
  let locale = deps.locale;
  let labels: SessionChartLabels = {
    ...DEFAULT_LABELS,
    ...deps.labels,
    close: deps.closeLabel ?? deps.labels?.close ?? DEFAULT_LABELS.close,
  };
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

  const updateSelectionClasses = (): void => {
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
    if (document.activeElement !== panel) panel.focus();
  };

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

  const applyDrag = (
    e: PointerEvent,
    chart: SessionChart,
    mode: 'move' | 'resize',
    panel: HTMLElement,
  ): void => {
    if (e.button !== 0 || !canPlace(chart)) return;
    e.preventDefault();
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
      if (mode === 'move') {
        preview({ x: Math.max(0, start.x + dx), y: Math.max(0, start.y + dy) });
      } else {
        preview({ w: Math.max(MIN_W, start.w + dx), h: Math.max(MIN_H, start.h + dy) });
      }
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
    if (
      target instanceof HTMLElement &&
      target.closest('button, input, select, textarea, option, [contenteditable="true"]')
    ) {
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
    placeChart(chart, {
      x: Math.max(0, current.x + dx),
      y: Math.max(0, current.y + dy),
    });
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
    charts.forEach((chart, index) => {
      const drawable = isDrawable(chart);
      if (drawable) visibleChartIds.add(chart.id);
      const position = sessionPosition(store, chart.nativeAnchor);
      const panel = document.createElement('section');
      panel.className = drawable ? 'fc-chart' : 'fc-chart fc-chart--unsupported';
      panel.dataset.chartId = chart.id;
      panel.tabIndex = 0;
      panel.setAttribute('role', 'group');
      panel.setAttribute('aria-roledescription', 'chart');
      panel.setAttribute('aria-selected', chart.id === selectedId ? 'true' : 'false');
      panel.setAttribute(
        'aria-keyshortcuts',
        canPlace(chart)
          ? 'ArrowLeft ArrowRight ArrowUp ArrowDown Shift+ArrowLeft Shift+ArrowRight Shift+ArrowUp Shift+ArrowDown Delete Backspace'
          : 'Delete Backspace',
      );
      panel.setAttribute('aria-label', chartLabel(chart, labels));
      panel.style.left = `${position.x}px`;
      panel.style.top = `${position.y}px`;
      panel.style.width = `${position.w}px`;
      panel.style.height = `${position.h}px`;
      panel.style.zIndex = chart.id === selectedId ? String(charts.length + 1) : String(index + 1);
      panel.classList.toggle('fc-chart--selected', chart.id === selectedId);
      panel.addEventListener('focus', () => {
        selectedId = chart.id;
        updateSelectionClasses();
      });
      panel.addEventListener('pointerdown', () => selectChart(chart.id, panel));
      panel.addEventListener('keydown', (e) => applyKeyboard(e, chart, panel));

      const header = document.createElement('div');
      header.className = 'fc-chart__header';
      header.addEventListener('pointerdown', (e) => {
        if (
          e.target instanceof Element &&
          e.target.closest('button, input, select, textarea, option, [contenteditable="true"]')
        ) {
          return;
        }
        applyDrag(e, chart, 'move', panel);
      });
      const title = document.createElement('div');
      title.className = 'fc-chart__title';
      title.textContent = chartLabel(chart, labels);
      const close = createSessionChartCloseButton(labels.close);
      close.addEventListener('click', () => removeChart(chart));

      const svg = document.createElementNS(SVG_NS, 'svg');
      svg.classList.add('fc-chart__plot');
      svg.setAttribute('role', 'img');
      svg.setAttribute('aria-label', title.textContent);
      const width = Math.max(1, position.w - PLOT_INSET_W);
      const height = Math.max(1, position.h - PLOT_INSET_H);
      if (drawable) {
        const settingsButtons = document.createElement('div');
        settingsButtons.className = 'fc-chart__settings';
        const settingTabs: readonly {
          id: Parameters<SessionChartSettingsHandle['open']>[3];
          label: string;
        }[] = [
          { id: 'elements', label: labels.elementsTab ?? chartSettingsDefaults.elementsTab },
          { id: 'styles', label: labels.stylesTab ?? chartSettingsDefaults.stylesTab },
          { id: 'filters', label: labels.filtersTab ?? chartSettingsDefaults.filtersTab },
        ];
        for (const item of settingTabs) {
          if (!item.id) continue;
          const button = createInteractionButton({
            className: 'fc-chart__settings-button',
            ariaLabel: item.label,
            text: item.label,
          });
          button.dataset.chartSettingsTab = item.id;
          button.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            settings.open(chart, chart.data, button, item.id);
          });
          settingsButtons.appendChild(button);
        }
        header.append(title, settingsButtons, close);
        renderChartScene(
          svg,
          layoutChart({
            // The overlay header owns the title; standalone scenes render their own.
            spec: { ...chart, title: undefined },
            data: chart.data,
            width,
            height,
            ...(locale === undefined ? {} : { locale }),
            formatValue,
            axisTitle: labels.axisTitle,
            // Series without a colour of their own take Excel's automatic colours in the workbook theme.
            ...themeAccents(),
          }),
        );
      } else {
        header.append(title, close);
        // The header already shows the title, so the frame carries only the message.
        renderChartScene(
          svg,
          layoutUnsupportedChart({ message: labels.unsupportedChart, width, height }),
        );
      }
      panel.append(header, svg);
      if (canPlace(chart)) {
        const resize = document.createElement('div');
        resize.className = 'fc-chart__resize';
        resize.setAttribute('role', 'separator');
        resize.setAttribute('aria-label', labels.resize);
        resize.addEventListener('pointerdown', (e) => applyDrag(e, chart, 'resize', panel));
        panel.appendChild(resize);
      }
      root.appendChild(panel);
      if (drawable) settings.sync(chart, chart.data, panel);
    });
    if (settings.isOpen() && !visibleChartIds.has(settings.chartId() ?? '')) {
      settings.close(true);
    }
  };

  const unsubscribe = store.subscribe(render);
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
      unsubscribeController();
      settings.detach();
      root.remove();
    },
  };
}
