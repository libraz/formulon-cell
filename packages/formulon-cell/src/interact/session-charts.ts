import { type SessionChartSeriesPoint, sessionChartSeries } from '../commands/session-chart.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import { mutators } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface SessionChartsHandle {
  refresh(): void;
  setLabels(labels: Partial<SessionChartLabels>): void;
  detach(): void;
}

export interface SessionChartLabels {
  close: string;
  resize: string;
  columnChart: string;
  lineChart: string;
}

const SVG_NS = 'http://www.w3.org/2000/svg';
const MIN_W = 240;
const MIN_H = 160;
const KEYBOARD_STEP = 8;

const DEFAULT_LABELS: SessionChartLabels = {
  close: 'Close',
  resize: 'Resize chart',
  columnChart: 'Column chart',
  lineChart: 'Line chart',
};

function appendSvg<K extends keyof SVGElementTagNameMap>(
  parent: Element,
  tag: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] {
  const el = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, String(value));
  parent.appendChild(el);
  return el;
}

function renderChartSvg(
  svg: SVGSVGElement,
  chart: SessionChart,
  series: readonly SessionChartSeriesPoint[],
): void {
  svg.replaceChildren();
  const w = 320;
  const h = 132;
  const pad = 20;
  appendSvg(svg, 'rect', { x: 0, y: 0, width: w, height: h, rx: 6, fill: 'transparent' });
  if (series.length === 0) return;
  const values = series.map((s) => s.value);
  const min = Math.min(0, ...values);
  const max = Math.max(0, ...values);
  const span = max - min || 1;
  const xFor = (i: number): number =>
    series.length === 1 ? w / 2 : pad + (i * (w - pad * 2)) / (series.length - 1);
  const yFor = (v: number): number => h - pad - ((v - min) / span) * (h - pad * 2);
  const zeroY = yFor(0);
  appendSvg(svg, 'line', {
    x1: pad,
    x2: w - pad,
    y1: zeroY,
    y2: zeroY,
    stroke: 'currentColor',
    'stroke-opacity': 0.24,
    'stroke-width': 1,
  });

  const color = chart.color ?? '#0f6cbd';
  if (chart.kind === 'line') {
    const points = series.map((s, i) => `${xFor(i)},${yFor(s.value)}`).join(' ');
    appendSvg(svg, 'polyline', {
      points,
      fill: 'none',
      stroke: color,
      'stroke-width': 2.5,
      'stroke-linejoin': 'round',
      'stroke-linecap': 'round',
    });
    for (let i = 0; i < series.length; i += 1) {
      appendSvg(svg, 'circle', { cx: xFor(i), cy: yFor(series[i]?.value ?? 0), r: 3, fill: color });
    }
    return;
  }

  const slot = (w - pad * 2) / Math.max(1, series.length);
  const bw = Math.max(6, Math.min(28, slot * 0.58));
  for (let i = 0; i < series.length; i += 1) {
    const v = series[i]?.value ?? 0;
    const y = yFor(v);
    appendSvg(svg, 'rect', {
      x: pad + i * slot + (slot - bw) / 2,
      y: Math.min(y, zeroY),
      width: bw,
      height: Math.max(1, Math.abs(zeroY - y)),
      rx: 2,
      fill: v < 0 ? '#d13438' : color,
    });
  }
}

export function attachSessionCharts(deps: {
  host: HTMLElement;
  store: SpreadsheetStore;
  closeLabel?: string;
  labels?: Partial<SessionChartLabels>;
}): SessionChartsHandle {
  const { host, store } = deps;
  let labels: SessionChartLabels = {
    ...DEFAULT_LABELS,
    ...deps.labels,
    close: deps.closeLabel ?? deps.labels?.close ?? DEFAULT_LABELS.close,
  };
  const root = document.createElement('div');
  root.className = 'fc-charts';
  host.appendChild(root);
  inheritHostTokens(host, root);
  let selectedId: string | null = null;

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
    const onMove = (move: PointerEvent): void => {
      const dx = move.clientX - startX;
      const dy = move.clientY - startY;
      if (mode === 'move') {
        mutators.updateChart(store, chart.id, {
          x: Math.max(0, startLeft + dx),
          y: Math.max(0, startTop + dy),
        });
      } else {
        mutators.updateChart(store, chart.id, {
          w: Math.max(MIN_W, startW + dx),
          h: Math.max(MIN_H, startH + dy),
        });
      }
    };
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp, { once: true });
  };

  const applyKeyboard = (e: KeyboardEvent, chart: SessionChart, panel: HTMLElement): void => {
    if (e.defaultPrevented) return;
    selectChart(chart.id, panel);
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      mutators.removeChart(store, chart.id);
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
      mutators.updateChart(store, chart.id, {
        w: Math.max(MIN_W, ((chart.w ?? panel.offsetWidth) || 360) + dx),
        h: Math.max(MIN_H, ((chart.h ?? panel.offsetHeight) || 220) + dy),
      });
      return;
    }

    mutators.updateChart(store, chart.id, {
      x: Math.max(0, (chart.x ?? panel.offsetLeft) + dx),
      y: Math.max(0, (chart.y ?? panel.offsetTop) + dy),
    });
  };

  const render = (): void => {
    const state = store.getState();
    root.replaceChildren();
    state.charts.charts
      .filter((chart) => chart.source.sheet === state.data.sheetIndex)
      .forEach((chart, idx) => {
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
        panel.setAttribute(
          'aria-label',
          chart.title ?? (chart.kind === 'line' ? labels.lineChart : labels.columnChart),
        );
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
          if (e.target === close) return;
          applyDrag(e, chart, 'move', panel);
        });
        const title = document.createElement('div');
        title.className = 'fc-chart__title';
        title.textContent =
          chart.title ?? (chart.kind === 'line' ? labels.lineChart : labels.columnChart);
        const close = document.createElement('button');
        close.type = 'button';
        close.className = 'fc-chart__close';
        close.setAttribute('aria-label', labels.close);
        close.textContent = '×';
        close.addEventListener('click', () => {
          mutators.removeChart(store, chart.id);
          host.focus({ preventScroll: true });
        });
        header.append(title, close);

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.classList.add('fc-chart__plot');
        svg.setAttribute('viewBox', '0 0 320 132');
        svg.setAttribute('role', 'img');
        svg.setAttribute('aria-label', title.textContent);
        renderChartSvg(svg, chart, sessionChartSeries(state, chart));
        const resize = document.createElement('div');
        resize.className = 'fc-chart__resize';
        resize.setAttribute('role', 'separator');
        resize.setAttribute('aria-label', labels.resize);
        resize.addEventListener('pointerdown', (e) => applyDrag(e, chart, 'resize', panel));
        panel.append(header, svg, resize);
        root.appendChild(panel);
      });
  };

  const unsubscribe = store.subscribe(render);
  render();

  return {
    refresh: render,
    setLabels(next) {
      labels = { ...labels, ...next };
      render();
    },
    detach() {
      unsubscribe();
      root.remove();
    },
  };
}
