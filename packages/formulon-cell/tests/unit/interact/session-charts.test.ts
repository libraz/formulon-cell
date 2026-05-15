import { describe, expect, it } from 'vitest';
import { attachSessionCharts } from '../../../src/interact/session-charts.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const firePointer = (
  target: EventTarget,
  type: string,
  init: PointerEventInit = {},
): PointerEvent => {
  const e = new PointerEvent(type, {
    bubbles: true,
    pointerId: 1,
    button: 0,
    clientX: 0,
    clientY: 0,
    ...init,
  });
  target.dispatchEvent(e);
  return e;
};

const fireKey = (target: EventTarget, key: string, init: KeyboardEventInit = {}): KeyboardEvent => {
  const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key, ...init });
  target.dispatchEvent(e);
  return e;
};

describe('attachSessionCharts', () => {
  it('renders session chart overlays and removes them from the store', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 10 });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'number', value: 20 });
    mutators.upsertChart(store, {
      id: 'chart-1',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      title: 'Revenue',
    });

    const handle = attachSessionCharts({ host, store, closeLabel: '閉じる' });

    expect(host.querySelector('.fc-chart__title')?.textContent).toBe('Revenue');
    expect(host.querySelector('svg rect[fill="#0f6cbd"]')).toBeTruthy();
    host.tabIndex = -1;

    const close = host.querySelector<HTMLButtonElement>('.fc-chart__close');
    expect(close?.getAttribute('aria-label')).toBe('閉じる');
    close?.click();
    expect(store.getState().charts.charts).toHaveLength(0);
    expect(host.querySelector('.fc-chart')).toBeNull();
    expect(document.activeElement).toBe(host);

    handle.detach();
  });

  it('uses localized fallback labels for untitled charts', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'chart-1',
      kind: 'line',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });

    const handle = attachSessionCharts({
      host,
      store,
      labels: {
        close: '閉じる',
        resize: 'グラフのサイズ変更',
        columnChart: '縦棒グラフ',
        lineChart: '折れ線グラフ',
      },
    });

    expect(host.querySelector('.fc-chart__title')?.textContent).toBe('折れ線グラフ');
    expect(host.querySelector('.fc-chart')?.getAttribute('aria-label')).toBe('折れ線グラフ');
    expect(host.querySelector('.fc-chart__resize')?.getAttribute('aria-label')).toBe(
      'グラフのサイズ変更',
    );

    handle.setLabels({ lineChart: '線グラフ', resize: 'サイズ変更' });
    expect(host.querySelector('.fc-chart__title')?.textContent).toBe('線グラフ');
    expect(host.querySelector('.fc-chart__resize')?.getAttribute('aria-label')).toBe('サイズ変更');

    handle.detach();
  });

  it('moves and resizes chart overlays through pointer interactions', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'chart-1',
      kind: 'line',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      title: 'Trend',
      x: 20,
      y: 30,
      w: 320,
      h: 180,
    });

    const handle = attachSessionCharts({ host, store });
    const header = host.querySelector<HTMLElement>('.fc-chart__header');
    expect(header).toBeTruthy();
    if (!header) throw new Error('missing chart header');
    firePointer(header, 'pointerdown', { clientX: 10, clientY: 10 });
    firePointer(window, 'pointermove', { clientX: 45, clientY: 60 });
    firePointer(window, 'pointerup', { clientX: 45, clientY: 60 });
    expect(store.getState().charts.charts[0]).toMatchObject({ x: 55, y: 80 });

    const resize = host.querySelector<HTMLElement>('.fc-chart__resize');
    expect(resize).toBeTruthy();
    if (!resize) throw new Error('missing chart resize handle');
    firePointer(resize, 'pointerdown', { clientX: 100, clientY: 100 });
    firePointer(window, 'pointermove', { clientX: 180, clientY: 145 });
    firePointer(window, 'pointerup', { clientX: 180, clientY: 145 });
    expect(store.getState().charts.charts[0]).toMatchObject({ w: 400, h: 225 });

    handle.detach();
  });

  it('marks the active chart overlay as selected and brings it forward', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'chart-1',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      title: 'First',
    });
    mutators.upsertChart(store, {
      id: 'chart-2',
      kind: 'line',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      title: 'Second',
    });

    const handle = attachSessionCharts({ host, store });
    const charts = Array.from(host.querySelectorAll<HTMLElement>('.fc-chart'));
    expect(charts).toHaveLength(2);
    const first = charts[0];
    const second = charts[1];
    if (!first || !second) throw new Error('missing chart overlays');

    firePointer(second, 'pointerdown');
    expect(second.classList.contains('fc-chart--selected')).toBe(true);
    expect(second.getAttribute('aria-selected')).toBe('true');
    expect(second.style.zIndex).toBe('2');
    expect(first.getAttribute('aria-selected')).toBe('false');

    first.focus();
    expect(first.classList.contains('fc-chart--selected')).toBe(true);
    expect(first.style.zIndex).toBe('2');
    expect(second.getAttribute('aria-selected')).toBe('false');

    handle.detach();
  });

  it('moves, resizes, and removes focused chart overlays through keyboard interactions', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'chart-1',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      title: 'Accessible chart',
      x: 20,
      y: 30,
      w: 320,
      h: 180,
    });

    const handle = attachSessionCharts({ host, store });
    const chart = host.querySelector<HTMLElement>('.fc-chart');
    expect(chart?.tabIndex).toBe(0);
    expect(chart?.getAttribute('aria-label')).toBe('Accessible chart');
    expect(chart?.getAttribute('aria-roledescription')).toBe('chart');
    expect(chart?.getAttribute('aria-keyshortcuts')).toContain('Shift+ArrowDown');
    if (!chart) throw new Error('missing chart');

    const move = fireKey(chart, 'ArrowRight');
    expect(move.defaultPrevented).toBe(true);
    expect(store.getState().charts.charts[0]).toMatchObject({ x: 28, y: 30 });

    const resize = fireKey(host.querySelector<HTMLElement>('.fc-chart') ?? chart, 'ArrowDown', {
      shiftKey: true,
    });
    expect(resize.defaultPrevented).toBe(true);
    expect(store.getState().charts.charts[0]).toMatchObject({ w: 320, h: 188 });

    fireKey(host.querySelector<HTMLElement>('.fc-chart') ?? chart, 'Delete');
    expect(store.getState().charts.charts).toHaveLength(0);
    expect(host.querySelector('.fc-chart')).toBeNull();
    expect(document.activeElement).toBe(host);

    handle.detach();
  });
});
