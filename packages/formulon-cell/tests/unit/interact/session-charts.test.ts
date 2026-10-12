import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { attachSessionCharts } from '../../../src/interact/session-charts.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

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

  it('keeps bar and column marks inside padded slots for sparse and dense series', () => {
    for (const kind of ['bar', 'column'] as const) {
      for (const count of [1, 2, 16, 100]) {
        const host = document.createElement('div');
        document.body.appendChild(host);
        const store = createSpreadsheetStore();
        for (let index = 0; index < count; index += 1) {
          const row = kind === 'bar' ? index : 0;
          const col = kind === 'bar' ? 0 : index;
          mutators.setCell(store, { sheet: 0, row, col }, { kind: 'number', value: index + 1 });
        }
        mutators.upsertChart(store, {
          id: `${kind}-chart-${count}`,
          kind,
          source: {
            sheet: 0,
            r0: 0,
            c0: 0,
            r1: kind === 'bar' ? count - 1 : 0,
            c1: kind === 'bar' ? 0 : count - 1,
          },
        });

        const handle = attachSessionCharts({ host, store });
        const marks = Array.from(
          host.querySelectorAll<SVGRectElement>('.fc-chart__plot rect[fill="#0f6cbd"]'),
        )
          .map((mark) => ({
            start: Number(mark.getAttribute(kind === 'bar' ? 'y' : 'x')),
            size: Number(mark.getAttribute(kind === 'bar' ? 'height' : 'width')),
          }))
          .sort((a, b) => a.start - b.start);
        expect(marks).toHaveLength(count);
        for (const mark of marks) {
          expect(mark.start).toBeGreaterThanOrEqual(20);
          expect(mark.start + mark.size).toBeLessThanOrEqual(kind === 'bar' ? 112 : 300);
          expect(mark.size).toBeGreaterThan(0);
        }
        for (let index = 1; index < marks.length; index += 1) {
          const previous = marks[index - 1];
          const current = marks[index];
          if (!previous || !current) throw new Error('chart mark missing from measured series');
          expect(current.start).toBeGreaterThanOrEqual(previous.start + previous.size - 1e-9);
        }
        handle.detach();
        host.remove();
      }
    }
  });

  it('routes chart object edits through the protection guard', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'protected-chart',
      kind: 'line',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      x: 20,
      y: 30,
      w: 320,
      h: 180,
    });

    const handle = attachSessionCharts({ host, store });
    const chart = host.querySelector<HTMLElement>('.fc-chart');
    const header = host.querySelector<HTMLElement>('.fc-chart__header');
    const close = host.querySelector<HTMLButtonElement>('.fc-chart__close');
    if (!chart || !header || !close) throw new Error('missing chart controls');
    setProtectedSheet(store, 0, true);
    const before = structuredClone(store.getState().charts.charts[0]);

    firePointer(header, 'pointerdown', { clientX: 10, clientY: 10 });
    firePointer(window, 'pointermove', { clientX: 45, clientY: 60 });
    firePointer(window, 'pointerup', { clientX: 45, clientY: 60 });
    fireKey(chart, 'ArrowRight');
    fireKey(chart, 'Delete');
    close.click();

    expect(store.getState().charts.charts).toHaveLength(1);
    expect(store.getState().charts.charts[0]).toEqual(before);
    setProtectedSheet(store, 0, false);
    fireKey(chart, 'ArrowRight');
    expect(store.getState().charts.charts[0]?.x).toBe(28);
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

  it('keeps chart close button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/session-charts.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('const createSessionChartCloseButton');
    expect(source).toContain('const close = createSessionChartCloseButton(labels.close)');
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps session charts close to Excel 365 desktop object geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/overlays/quick-analysis-and-charts.css'),
      'utf8',
    );
    const chartCss = css.slice(css.indexOf('.fc-chart {'));

    expect(chartCss).toMatch(/\.fc-chart\s*\{[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow:/);
    expect(chartCss).toMatch(
      /\.fc-chart:focus-visible\s*\{[\s\S]*?0 0 0 1px var\(--fc-accent, Highlight\)/,
    );
    expect(chartCss).toMatch(
      /\.fc-chart--selected::after\s*\{[\s\S]*?inset: -2px;[\s\S]*?border: 1px solid var\(--fc-accent, Highlight\);[\s\S]*?border-radius: 2px;/,
    );
    expect(chartCss).toMatch(
      /\.fc-chart__header\s*\{[\s\S]*?min-height: 30px;[\s\S]*?padding: 0 6px 0 10px;/,
    );
    expect(chartCss).toMatch(
      /\.fc-chart__close\s*\{[\s\S]*?width: 22px;[\s\S]*?height: 22px;[\s\S]*?border-radius: 2px;/,
    );
    expect(chartCss).toMatch(
      /\.fc-chart__resize\s*\{[\s\S]*?right: 2px;[\s\S]*?bottom: 2px;[\s\S]*?width: 12px;[\s\S]*?height: 12px;/,
    );
    expect(chartCss).toMatch(
      /\.fc-chart__resize::before\s*\{[\s\S]*?width: 7px;[\s\S]*?height: 7px;[\s\S]*?border-right: 1px solid/,
    );
    expect(chartCss).not.toContain('border-radius: var(--fc-radius-md, 6px);');
    expect(chartCss).not.toContain('border: 1px dashed var(--fc-accent');
  });
});
