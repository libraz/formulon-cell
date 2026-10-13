import { layoutChart } from '../charts/layout.js';
import type { ChartPrimitive, ChartScene } from '../charts/types.js';
import type { History } from '../commands/history.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { isSheetProtected } from '../commands/protection.js';
import {
  clearSessionChart,
  sessionChartData,
  updateSessionChart,
} from '../commands/session-chart.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import { createInteractionButton } from './chip-button.js';
import { formatSessionChartValue } from './session-chart-format.js';
import {
  type ChartValueFormatter,
  chartSettingsDefaults,
  createSessionChartSettings,
  type SessionChartSettingsHandle,
  type SessionChartSettingsLabels,
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
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const MIN_W = 240;
const MIN_H = 160;
const KEYBOARD_STEP = 8;
const SESSION_CHART_PALETTE = ['#0f6cbd', '#ed7d31', '#70ad47', '#ffc000', '#5b9bd5', '#a5a5a5'];

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

const chartLabel = (chart: SessionChart, labels: SessionChartLabels): string => {
  if (chart.title) return chart.title;
  if (chart.kind === 'bar') return labels.barChart;
  if (chart.kind === 'line') return labels.lineChart;
  if (chart.kind === 'area') return labels.areaChart;
  if (chart.kind === 'pie') return labels.pieChart;
  if (chart.kind === 'doughnut') return labels.doughnutChart;
  if (chart.kind === 'scatter') return labels.scatterChart;
  return labels.columnChart;
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
  closeLabel?: string;
  labels?: Partial<SessionChartLabels>;
  history?: History | null;
  locale?: string;
  formatValue?: ChartValueFormatter;
}): SessionChartsHandle {
  const { host, store, history = null } = deps;
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

  const canEditCharts = (chart: SessionChart): boolean => {
    if (isSheetProtected(store.getState(), chart.source.sheet)) return false;
    const controller = interactionControllerFor(store);
    if (!controller) return true;
    return controller.canExecute({
      operation: 'object',
      origin: 'instanceApi',
      effects: [
        {
          kind: 'cells',
          cells: [
            {
              sheet: chart.source.sheet,
              row: chart.source.r0,
              col: chart.source.c0,
            },
          ],
        },
      ],
    }).allowed;
  };

  settings = createSessionChartSettings({
    host,
    store,
    history,
    locale,
    formatValue,
    labels: { ...SETTINGS_TAB_LABELS, ...deps.labels },
    canEdit: canEditCharts,
    readData: (draftChart) => sessionChartData(store.getState(), draftChart),
    onApply: (id, patch, patchHistory) => {
      updateSessionChart(
        store,
        id,
        patch as Parameters<typeof updateSessionChart>[2],
        patchHistory,
      );
    },
  });

  const updateSelectionClasses = (): void => {
    const panels = Array.from(root.querySelectorAll<HTMLElement>('.fc-chart'));
    for (const panel of panels) {
      const selected = panel.dataset.chartId === selectedId;
      panel.classList.toggle('fc-chart--selected', selected);
      panel.setAttribute('aria-selected', selected ? 'true' : 'false');
      panel.style.zIndex = selected ? '2' : '1';
    }
  };

  const selectChart = (chartId: string, panel: HTMLElement): void => {
    selectedId = chartId;
    updateSelectionClasses();
    if (document.activeElement !== panel) panel.focus();
  };

  const applyDrag = (
    e: PointerEvent,
    chart: SessionChart,
    mode: 'move' | 'resize',
    panel: HTMLElement,
  ): void => {
    if (e.button !== 0) return;
    e.preventDefault();
    const startX = e.clientX;
    const startY = e.clientY;
    const startLeft = chart.x ?? panel.offsetLeft;
    const startTop = chart.y ?? panel.offsetTop;
    const startW = chart.w ?? (panel.offsetWidth || 360);
    const startH = chart.h ?? (panel.offsetHeight || 220);
    let latest: { x?: number; y?: number; w?: number; h?: number } = {};
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
        preview({
          x: Math.max(0, startLeft + dx),
          y: Math.max(0, startTop + dy),
        });
      } else {
        preview({
          w: Math.max(MIN_W, startW + dx),
          h: Math.max(MIN_H, startH + dy),
        });
      }
    };
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      const committed = updateSessionChart(store, chart.id, latest, history);
      if (!committed) render();
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
      clearSessionChart(store, chart.id, history);
      host.focus({ preventScroll: true });
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

    const [dx, dy] = delta;
    if (e.shiftKey) {
      updateSessionChart(
        store,
        chart.id,
        {
          w: Math.max(MIN_W, ((chart.w ?? panel.offsetWidth) || 360) + dx),
          h: Math.max(MIN_H, ((chart.h ?? panel.offsetHeight) || 220) + dy),
        },
        history,
      );
      return;
    }

    updateSessionChart(
      store,
      chart.id,
      {
        x: Math.max(0, (chart.x ?? panel.offsetLeft) + dx),
        y: Math.max(0, (chart.y ?? panel.offsetTop) + dy),
      },
      history,
    );
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
    state.charts.charts
      .filter((chart) => chart.source.sheet === state.data.sheetIndex)
      .forEach((chart, idx) => {
        visibleChartIds.add(chart.id);
        const panel = document.createElement('section');
        panel.className = 'fc-chart';
        panel.dataset.chartId = chart.id;
        panel.tabIndex = 0;
        panel.setAttribute('role', 'group');
        panel.setAttribute('aria-roledescription', 'chart');
        panel.setAttribute('aria-selected', chart.id === selectedId ? 'true' : 'false');
        panel.setAttribute(
          'aria-keyshortcuts',
          'ArrowLeft ArrowRight ArrowUp ArrowDown Shift+ArrowLeft Shift+ArrowRight Shift+ArrowUp Shift+ArrowDown Delete Backspace',
        );
        panel.setAttribute('aria-label', chartLabel(chart, labels));
        panel.style.left = `${chart.x ?? 320 + idx * 24}px`;
        panel.style.top = `${chart.y ?? 72 + idx * 24}px`;
        panel.style.width = `${chart.w ?? 360}px`;
        panel.style.height = `${chart.h ?? 220}px`;
        panel.style.zIndex = chart.id === selectedId ? '2' : '1';
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
            settings.open(chart, data, button, item.id);
          });
          settingsButtons.appendChild(button);
        }
        const close = createSessionChartCloseButton(labels.close);
        close.addEventListener('click', () => {
          clearSessionChart(store, chart.id, history);
          host.focus({ preventScroll: true });
        });
        header.append(title, settingsButtons, close);

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.classList.add('fc-chart__plot');
        svg.setAttribute('viewBox', '0 0 320 132');
        svg.setAttribute('role', 'img');
        svg.setAttribute('aria-label', title.textContent);
        const data = sessionChartData(state, chart);
        const spec = {
          ...chart,
          // The overlay header owns the title; standalone scenes render their own.
          title: undefined,
          // Preserve the legacy session-chart colour when a host supplied one
          // while allowing the pure chart module to own palette selection.
          palette:
            chart.palette && chart.palette.length > 0
              ? chart.palette
              : [
                  chart.color ?? SESSION_CHART_PALETTE[0] ?? '#0f6cbd',
                  ...SESSION_CHART_PALETTE.slice(1),
                ],
        };
        const layoutInput = {
          spec,
          data,
          width: 320,
          height: 132,
          locale,
          formatValue,
        };
        renderChartScene(svg, layoutChart(layoutInput as Parameters<typeof layoutChart>[0]));
        const resize = document.createElement('div');
        resize.className = 'fc-chart__resize';
        resize.setAttribute('role', 'separator');
        resize.setAttribute('aria-label', labels.resize);
        resize.addEventListener('pointerdown', (e) => applyDrag(e, chart, 'resize', panel));
        panel.append(header, svg, resize);
        root.appendChild(panel);
        settings.sync(chart, data, panel);
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
