import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { createSessionChart, updateSessionChart } from '../../../src/commands/session-chart.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatSessionChartValue } from '../../../src/interact/session-chart-format.js';
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

const chartAxisFieldset = (surface: HTMLElement, axisLabel: string): HTMLElement => {
  const fieldset = Array.from(surface.querySelectorAll<HTMLElement>('fieldset')).find(
    (candidate) => candidate.querySelector('legend')?.textContent === axisLabel,
  );
  if (!fieldset) throw new Error(`missing axis fieldset: ${axisLabel}`);
  return fieldset;
};

const chartAxisCrossingSelect = (surface: HTMLElement, axisLabel: string): HTMLSelectElement => {
  const input = chartAxisFieldset(surface, axisLabel).querySelector<HTMLSelectElement>(
    '.fc-chart-settings__crossing-mode select',
  );
  if (!input) throw new Error(`missing crossing select: ${axisLabel}`);
  return input;
};

const chartAxisCrossingCustom = (surface: HTMLElement, axisLabel: string): HTMLInputElement => {
  const fieldset = chartAxisFieldset(surface, axisLabel);
  const field = Array.from(
    fieldset.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
  ).find((candidate) => candidate.textContent?.includes('Custom value'));
  const input = field?.querySelector<HTMLInputElement>('input');
  if (!input) throw new Error(`missing custom crossing input: ${axisLabel}`);
  return input;
};

describe('attachSessionCharts', () => {
  it('uses the renderer auto-legend default for one series and allows an explicit legend', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    [
      ['Month', 'Sales'],
      ['Jan', 10],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'single-series-legend', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store, history });
    const openElements = (): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };
    const legendCheckbox = (surface: HTMLElement): HTMLInputElement => {
      const row = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__check'),
      ).find((candidate) => candidate.textContent?.trim() === 'Legend');
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error('missing legend checkbox');
      return input;
    };

    const initial = openElements();
    expect(legendCheckbox(initial).checked).toBe(false);
    initial.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(history.canUndo()).toBe(false);

    const changed = openElements();
    const checkbox = legendCheckbox(changed);
    checkbox.checked = true;
    checkbox.dispatchEvent(new Event('change', { bubbles: true }));
    changed.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.legend?.visible).toBe(true);
    expect(host.querySelectorAll('[data-chart-role="legend"]')).not.toHaveLength(0);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.legend).toBeUndefined();
    expect(host.querySelectorAll('[data-chart-role="legend"]')).toHaveLength(0);
    handle.detach();
    host.remove();
  });

  it('opens a non-modal settings surface and applies labels with one undoable update', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'Month' });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'Sales' });
    mutators.setCell(store, { sheet: 0, row: 1, col: 0 }, { kind: 'text', value: 'Jan' });
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 10 });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'settings-chart', kind: 'column', title: 'Old title' },
    );
    const handle = attachSessionCharts({ host, store, history });
    const settingsButton = host.querySelector<HTMLButtonElement>(
      '.fc-chart__settings-button[data-chart-settings-tab="elements"]',
    );
    expect(settingsButton).toBeTruthy();
    settingsButton?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    expect(surface?.getAttribute('role')).toBe('dialog');
    expect(surface?.getAttribute('aria-modal')).toBe('false');
    expect(surface?.parentElement).toBe(host);
    expect(surface?.querySelectorAll('[role="tab"]')).toHaveLength(3);
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    expect(title?.value).toBe('Old title');
    if (!title) throw new Error('missing chart title input');
    title.value = 'New title';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();

    expect(store.getState().charts.charts[0]?.title).toBe('New title');
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.title).toBe('Old title');
    expect(history.redo()).toBe(true);
    expect(store.getState().charts.charts[0]?.title).toBe('New title');
    handle.detach();
    host.remove();
  });

  it('keeps filter drafts outside the chart root and commits selected series/categories', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const table = [
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
      ['Feb', 30, 40],
    ];
    table.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      { id: 'filter-chart', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="filters"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    expect(surface?.closest('.fc-charts')).toBeNull();
    const filterGroups = surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group');
    expect(filterGroups).toHaveLength(2);
    const seriesChecks =
      filterGroups?.[0]?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
    const categoryChecks =
      filterGroups?.[1]?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
    expect(seriesChecks).toHaveLength(2);
    expect(categoryChecks).toHaveLength(2);
    seriesChecks?.[0]?.click();
    categoryChecks?.[1]?.click();
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.filter).toEqual({
      hiddenSeriesIds: ['series-0'],
      hiddenCategoryIds: ['category-1'],
    });
    expect(host.querySelectorAll('[data-chart-role="bar"]')).toHaveLength(1);
    handle.detach();
    host.remove();
  });

  it('keeps pie variation stable for title-only and no-op Apply', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    [
      ['A', 'B'],
      [1, 2],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'pie-stability', kind: 'pie' },
    );
    const handle = attachSessionCharts({ host, store, history });
    const fills = (): string[] =>
      Array.from(host.querySelectorAll<SVGElement>('[data-chart-role="slice"]')).map(
        (slice) => slice.getAttribute('fill') ?? '',
      );
    const before = fills();
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing title input');
    title.value = 'Updated pie';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(fills()).toEqual(before);
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    host
      .querySelector<HTMLElement>('.fc-chart-settings')
      ?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')
      ?.click();
    expect(history.canUndo()).toBe(false);
    expect(fills()).toEqual(before);
    handle.detach();
    host.remove();
  });

  it('applies palette, series styles, point styles, and cancels drafts with focus restoration', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 1 });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'number', value: 2 });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      { id: 'style-chart', kind: 'line' },
    );
    const history = new History();
    const handle = attachSessionCharts({ host, store, history });
    expect(host.querySelector('[data-chart-role="line"]')?.getAttribute('data-series-id')).toBe(
      'series-0',
    );
    expect(host.querySelector('[data-chart-role="point"]')?.getAttribute('data-point-id')).toBe(
      'category-0',
    );
    const stylesButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-settings-tab="styles"]',
    );
    expect(stylesButton).toBeTruthy();
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    expect(
      surface?.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.textContent,
    ).toBe('Styles');
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(history.canUndo()).toBe(false);
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const styleGroups = surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group');
    const target = styleGroups?.[1];
    const seriesFill = target?.querySelector<HTMLInputElement>('input[type="color"]');
    expect(seriesFill).toBeTruthy();
    seriesFill?.setAttribute('value', '#ff0000');
    if (seriesFill) {
      seriesFill.value = '#ff0000';
      seriesFill.dispatchEvent(new Event('input', { bubbles: true }));
    }
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.seriesStyles?.[0]?.fill?.color).toBe('#ff0000');

    const reopenedStylesButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-settings-tab="styles"]',
    );
    reopenedStylesButton?.click();
    const reopened = host.querySelector<HTMLElement>('.fc-chart-settings');
    const reopenedSelects = reopened?.querySelectorAll<HTMLSelectElement>('select');
    const point = reopenedSelects?.[1];
    if (point) {
      point.value = 'category-0';
      point.dispatchEvent(new Event('change', { bubbles: true }));
    }
    const pointSurface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const pointGroups = pointSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group');
    const pointFill = pointGroups?.[1]?.querySelector<HTMLInputElement>('input[type="color"]');
    if (pointFill) {
      pointFill.value = '#00aa00';
      pointFill.dispatchEvent(new Event('input', { bubbles: true }));
    }
    pointSurface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.seriesStyles?.[0]?.points?.[0]?.fill?.color).toBe(
      '#00aa00',
    );

    const cancelStylesButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-settings-tab="styles"]',
    );
    cancelStylesButton?.click();
    const cancelSurface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const cancelTitle = cancelSurface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (cancelTitle) {
      cancelTitle.value = 'discarded';
      cancelTitle.dispatchEvent(new Event('input', { bubbles: true }));
    }
    cancelSurface?.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
    expect(store.getState().charts.charts[0]?.title).toBe('Line chart');
    expect(document.activeElement).toBe(cancelStylesButton);
    const elementsButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-settings-tab="elements"]',
    );
    elementsButton?.click();
    const escapeSurface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const escapeTitle = escapeSurface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (escapeTitle) {
      escapeTitle.value = 'escaped draft';
      escapeTitle.dispatchEvent(new Event('input', { bubbles: true }));
    }
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(store.getState().charts.charts[0]?.title).toBe('Line chart');
    expect(document.activeElement).toBe(elementsButton);
    handle.detach();
    host.remove();
  });

  it('shows effective palette and inherited point colors in style controls', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const table = [
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
      ['Feb', 30, 40],
    ];
    table.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      { id: 'effective-colors', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store });
    const firstBar = host.querySelector<SVGElement>(
      '[data-chart-role="bar"][data-series-index="0"]',
    );
    const secondBar = host.querySelector<SVGElement>(
      '[data-chart-role="bar"][data-series-index="1"]',
    );
    expect(firstBar?.getAttribute('fill')).toBe('#0f6cbd');
    expect(secondBar?.getAttribute('fill')).toBe('#ed7d31');
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    let surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    let groups = surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group');
    let target = groups?.[1];
    expect(target?.querySelector<HTMLInputElement>('input[type="color"]')?.value).toBe('#0f6cbd');
    const seriesSelect = target?.querySelectorAll<HTMLSelectElement>('select')[0];
    if (!seriesSelect) throw new Error('missing series selector');
    seriesSelect.value = 'series-1';
    seriesSelect.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    groups = surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group');
    target = groups?.[1];
    expect(target?.querySelector<HTMLInputElement>('input[type="color"]')?.value).toBe('#ed7d31');
    handle.detach();
    host.remove();

    const pieHost = document.createElement('div');
    document.body.appendChild(pieHost);
    const pieStore = createSpreadsheetStore();
    [
      ['A', 'B'],
      [1, 2],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          pieStore,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      pieStore,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'effective-pie-colors', kind: 'pie', seriesDirection: 'rows' },
    );
    const pieHandle = attachSessionCharts({ host: pieHost, store: pieStore });
    const secondSlice = pieHost.querySelector<SVGElement>(
      '[data-chart-role="slice"][data-point-index="1"]',
    );
    expect(secondSlice?.getAttribute('fill')).toBe('#ed7d31');
    pieHost.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const pieSurface = pieHost.querySelector<HTMLElement>('.fc-chart-settings');
    const pieTarget = pieSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const pointSelect = pieTarget?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (!pointSelect) throw new Error('missing point selector');
    pointSelect.value = 'category-1';
    pointSelect.dispatchEvent(new Event('change', { bubbles: true }));
    const pointSurface = pieHost.querySelector<HTMLElement>('.fc-chart-settings');
    const pointTarget = pointSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    expect(pointTarget?.querySelector<HTMLInputElement>('input[type="color"]')?.value).toBe(
      secondSlice?.getAttribute('fill'),
    );
    pieHandle.detach();
    pieHost.remove();
  });

  it('keeps point overrides sparse when editing inherited fill, line, and marker fields', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    [
      ['Jan', 'Feb'],
      [10, 20],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      {
        id: 'sparse-point-style',
        kind: 'line',
        seriesDirection: 'rows',
        seriesStyles: [
          {
            seriesId: 'series-0',
            fill: { color: '#ff0000', transparency: 25 },
            line: { color: '#00aa00', transparency: 30, width: 2 },
            marker: { visible: true, shape: 'circle', size: 8 },
          },
        ],
      },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    let surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    let target = surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const pointSelect = target?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (!pointSelect) throw new Error('missing point selector');
    pointSelect.value = 'category-0';
    pointSelect.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    target = surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const numericInputs = target?.querySelectorAll<HTMLInputElement>('input[type="number"]');
    const fillTransparency = numericInputs?.[0];
    const lineTransparency = numericInputs?.[1];
    const markerSize = numericInputs?.[3];
    if (!fillTransparency || !lineTransparency || !markerSize) {
      throw new Error('missing point style inputs');
    }
    fillTransparency.value = '40';
    fillTransparency.dispatchEvent(new Event('input', { bubbles: true }));
    lineTransparency.value = '50';
    lineTransparency.dispatchEvent(new Event('input', { bubbles: true }));
    markerSize.value = '12';
    markerSize.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    const style = store.getState().charts.charts[0]?.seriesStyles?.[0];
    expect(style?.fill).toEqual({ color: '#ff0000', transparency: 25 });
    expect(style?.line).toEqual({ color: '#00aa00', transparency: 30, width: 2 });
    expect(style?.marker).toEqual({ visible: true, shape: 'circle', size: 8 });
    expect(style?.points?.[0]).toEqual({
      pointId: 'category-0',
      fill: { transparency: 40 },
      line: { transparency: 50 },
      marker: { size: 12 },
    });
    handle.detach();
    host.remove();
  });

  it('keeps settings drafts through a locale update and removes listeners on detach', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    createSessionChart(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 }, { id: 'locale-chart' });
    const handle = attachSessionCharts({ host, store, locale: 'en-US' });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing title input');
    title.value = 'draft title';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    handle.setLabels(
      {
        settings: 'グラフ設定',
        elementsTab: '要素',
        stylesTab: 'スタイル',
        filtersTab: 'フィルター',
      },
      'ja-JP',
    );
    expect(surface?.lang).toBe('ja-JP');
    expect(surface?.querySelector('.fc-chart-settings__title')?.textContent).toBe('グラフ設定');
    expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.value).toBe(
      'draft title',
    );
    handle.detach();
    expect(host.querySelector('.fc-chart-settings')).toBeNull();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    host.remove();
  });

  it('reformats SVG data labels when the session locale changes', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 1234.5 });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 'locale-label-chart',
        kind: 'column',
        dataLabels: { visible: true, showValue: true },
      },
    );
    const handle = attachSessionCharts({
      host,
      store,
      locale: 'en-US',
      formatValue: formatSessionChartValue,
    });
    const label = (): string | null =>
      host.querySelector<SVGTextElement>('[data-chart-role="data-label"]')?.textContent ?? null;
    expect(label()).toBe('1,234.5');
    handle.setLabels({}, 'de-DE');
    expect(label()).toBe('1.234,5');
    handle.detach();
    host.remove();
  });

  it('uses Excel chart label positions and preserves an invalid current position', () => {
    const openPositionSelect = (
      kind: 'column' | 'line' | 'area' | 'pie',
      grouping?: 'clustered' | 'stacked',
      position?: 'outside-end',
    ): {
      host: HTMLDivElement;
      store: ReturnType<typeof createSpreadsheetStore>;
      handle: ReturnType<typeof attachSessionCharts>;
      select: HTMLSelectElement;
    } => {
      const host = document.createElement('div');
      document.body.appendChild(host);
      const store = createSpreadsheetStore();
      mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 10 });
      createSessionChart(
        store,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        {
          id: `${kind}-${grouping ?? 'default'}`,
          kind,
          grouping,
          dataLabels: { visible: true, ...(position ? { position } : {}) },
        },
      );
      const handle = attachSessionCharts({ host, store });
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const selects = host.querySelectorAll<HTMLSelectElement>('.fc-chart-settings select');
      const select = Array.from(selects).find((candidate) =>
        candidate.parentElement?.textContent?.includes('Position'),
      );
      if (!select) throw new Error('missing chart label position select');
      return { host, store, handle, select };
    };

    const pie = openPositionSelect('pie');
    expect(
      pie.select.querySelector<HTMLOptionElement>('option[value="inside-end"]')?.disabled,
    ).toBe(false);
    pie.select.value = 'inside-end';
    pie.select.dispatchEvent(new Event('change', { bubbles: true }));
    pie.host.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(pie.store.getState().charts.charts[0]?.dataLabels?.position).toBe('inside-end');
    pie.handle.detach();
    pie.host.remove();

    const stacked = openPositionSelect('column', 'stacked', 'outside-end');
    expect(
      stacked.select.querySelector<HTMLOptionElement>('option[value="outside-end"]')?.disabled,
    ).toBe(true);
    stacked.handle.detach();
    stacked.host.remove();

    const line = openPositionSelect('line');
    expect(line.select.querySelector('option[value="best-fit"]')?.textContent).toBe('Automatic');
    expect(['best-fit', 'right']).toContain(line.select.value);
    line.handle.detach();
    line.host.remove();

    const area = openPositionSelect('area');
    expect(area.select.disabled).toBe(true);
    expect(area.select.querySelector<HTMLOptionElement>('option[value="best-fit"]')?.disabled).toBe(
      true,
    );
    area.handle.detach();
    area.host.remove();
  });

  it('uses the built-in Excel formatter for direct chart attachment', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 1234.5 });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 'default-formatter-chart',
        kind: 'column',
        dataLabels: { visible: true, showValue: true, numberFormat: '#,##0.00' },
      },
    );
    const handle = attachSessionCharts({ host, store, locale: 'en-US' });
    expect(host.querySelector<SVGTextElement>('[data-chart-role="data-label"]')?.textContent).toBe(
      '1,234.50',
    );
    handle.detach();
    host.remove();
  });

  it('patches only changed data-label fields and preserves sparse API values', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 'sparse-labels',
        kind: 'column',
        dataLabels: { visible: true, showValue: true, color: '#123456', fontSize: 72 },
      },
      history,
    );
    const handle = attachSessionCharts({ host, store, history });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const dataLabelFields = Array.from(
      surface?.querySelectorAll<HTMLElement>('fieldset') ?? [],
    ).find((fieldset) => fieldset.querySelector('legend')?.textContent === 'Data labels');
    const dataLabelNumberFormat = Array.from(
      dataLabelFields?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    ).find((field) => field.textContent?.includes('Number format'));
    const numberFormat = dataLabelNumberFormat?.querySelector<HTMLInputElement>('input');
    if (!numberFormat) throw new Error('missing number format input');
    numberFormat.value = '#,##0.00';
    numberFormat.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.dataLabels).toEqual({
      visible: true,
      showValue: true,
      color: '#123456',
      fontSize: 72,
      numberFormat: '#,##0.00',
    });
    expect(history.canUndo()).toBe(true);
    handle.detach();
    host.remove();
  });

  it('keeps implicit value labels when another label flag is enabled', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    [
      ['Month', 'Sales'],
      ['Jan', 10],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'implicit-value-label', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const findCheckbox = (label: string): HTMLInputElement | undefined =>
      Array.from(surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [])
        .find((row) => row.textContent?.includes(label))
        ?.querySelector<HTMLInputElement>('input') ?? undefined;
    const showDataLabels = findCheckbox('Show data labels');
    const showCategoryName = findCheckbox('Category name');
    if (!showDataLabels || !showCategoryName) throw new Error('missing data-label controls');
    showDataLabels.checked = true;
    showDataLabels.dispatchEvent(new Event('change', { bubbles: true }));
    showCategoryName.checked = true;
    showCategoryName.dispatchEvent(new Event('change', { bubbles: true }));
    const dataLabelFields = Array.from(
      surface?.querySelectorAll<HTMLElement>('fieldset') ?? [],
    ).find((fieldset) => fieldset.querySelector('legend')?.textContent === 'Data labels');
    const numberFormatField = Array.from(
      dataLabelFields?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    ).find((field) => field.textContent?.includes('Number format'));
    const numberFormat = numberFormatField?.querySelector<HTMLInputElement>('input');
    if (!numberFormat) throw new Error('missing number format input');
    numberFormat.value = '#,##0.00';
    numberFormat.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.dataLabels?.showValue).toBe(true);
    expect(host.querySelector<SVGTextElement>('[data-chart-role="data-label"]')?.textContent).toBe(
      'Jan, 10.00',
    );
    handle.detach();
    host.remove();
  });

  it('rejects settings Apply on a protected chart source', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      { id: 'protected-settings', kind: 'column', title: 'Original' },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing title input');
    title.value = 'Blocked';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    setProtectedSheet(store, 0, true);
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.title).toBe('Original');
    handle.detach();
    host.remove();
  });

  it('updates open settings editability when protection changes', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      { id: 'live-protection', kind: 'column', title: 'Original' },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing title input');
    title.value = 'Draft';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    setProtectedSheet(store, 0, true);
    expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.disabled).toBe(true);
    expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.value).toBe('Draft');
    setProtectedSheet(store, 0, false);
    expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.disabled).toBe(false);
    expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.value).toBe('Draft');
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
    expect(store.getState().charts.charts[0]?.title).toBe('Original');
    handle.detach();
    host.remove();
  });

  it('updates open settings editability when interaction policy changes', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({ store, history, getWb: () => workbook });
    const unregister = registerInteractionController(store, controller);
    try {
      createSessionChart(
        store,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        { id: 'live-policy', kind: 'column', title: 'Original' },
        history,
      );
      const handle = attachSessionCharts({ host, store, history });
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
      const apply = surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply');
      if (!title || !apply) throw new Error('missing live policy settings controls');
      title.value = 'Draft';
      title.dispatchEvent(new Event('input', { bubbles: true }));
      expect(apply.disabled).toBe(false);

      controller.setPolicy(viewerPolicy());
      expect(surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.disabled).toBe(
        true,
      );
      expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.value).toBe('Draft');

      controller.setPolicy(undefined);
      expect(surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.disabled).toBe(
        false,
      );
      expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.value).toBe('Draft');
      handle.detach();
      expect(() => controller.setPolicy(viewerPolicy())).not.toThrow();
      expect(host.querySelector('.fc-chart-settings')).toBeNull();
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
    host.remove();
  });

  it('shows a read-only settings surface for viewer policy', async () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({ store, history, getWb: () => workbook });
    const unregister = registerInteractionController(store, controller);
    try {
      createSessionChart(
        store,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        { id: 'viewer-settings', kind: 'column', title: 'Original' },
        history,
      );
      const handle = attachSessionCharts({ host, store, history });
      controller.setPolicy(viewerPolicy());
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      expect(surface?.querySelector<HTMLInputElement>('input[type="text"]')?.disabled).toBe(true);
      expect(surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.disabled).toBe(
        true,
      );
      surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
      expect(store.getState().charts.charts[0]?.title).toBe('Original');
      handle.detach();
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
    host.remove();
  });

  it('applies gap width and overlap updates to the displayed columns', () => {
    const host = document.createElement('div');
    const store = createSpreadsheetStore();
    [
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
    ].forEach((row, r) => {
      row.forEach((value, col) => {
        mutators.setCell(
          store,
          { sheet: 0, row: r, col },
          typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 2 },
      {
        id: 'geometry',
        kind: 'column',
        seriesDirection: 'columns',
        gapWidth: 150,
        overlap: 0,
      },
    );
    const handle = attachSessionCharts({ host, store });
    try {
      const before = Array.from(host.querySelectorAll('[data-chart-role="bar"]'));
      expect(before).toHaveLength(2);
      const width = Number(before[0]?.getAttribute('width'));
      expect(width).toBeGreaterThan(0);
      expect(Number(before[1]?.getAttribute('x'))).toBeCloseTo(
        Number(before[0]?.getAttribute('x')) + width,
      );
      updateSessionChart(store, 'geometry', { gapWidth: 0, overlap: 100 });
      const after = Array.from(host.querySelectorAll('[data-chart-role="bar"]'));
      expect(after).toHaveLength(2);
      expect(Number(after[0]?.getAttribute('width'))).toBeCloseTo(width * 3.5);
      expect(after[0]?.getAttribute('x')).toBe(after[1]?.getAttribute('x'));
    } finally {
      handle.detach();
    }
  });

  it('edits axis fields sparsely, preserves unknown values, and accepts negative and zero values', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertChart(store, {
      id: 'axis-settings',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      xAxis: { title: 'Original X', customAxisValue: 'keep-x' } as never,
      yAxis: {
        min: -10,
        max: 10,
        majorUnit: 2,
        numberFormat: '0.0',
        majorGridlines: false,
        customAxisValue: 'keep-y',
      } as never,
    });
    const handle = attachSessionCharts({ host, store, history });
    const openElements = (): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };
    const fieldInput = (fieldset: HTMLElement, text: string): HTMLInputElement => {
      const field = Array.from(
        fieldset.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find(
        (candidate) =>
          !candidate.classList.contains('fc-chart-settings__crossing-mode') &&
          candidate.querySelector('.fc-chart-settings__label')?.textContent === text,
      );
      const input = field?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing axis field: ${text}`);
      return input;
    };
    const axisFieldset = (surface: HTMLElement, label: string): HTMLElement => {
      const fieldset = Array.from(surface.querySelectorAll<HTMLElement>('fieldset')).find(
        (candidate) => candidate.querySelector('legend')?.textContent === label,
      );
      if (!fieldset) throw new Error(`missing axis fieldset: ${label}`);
      return fieldset;
    };

    const surface = openElements();
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(history.canUndo()).toBe(false);

    const changedSurface = openElements();
    const xAxis = axisFieldset(changedSurface, 'X-axis');
    const yAxis = axisFieldset(changedSurface, 'Y-axis');
    const xTitle = fieldInput(xAxis, 'X-axis title');
    xTitle.value = 'Updated X';
    xTitle.dispatchEvent(new Event('input', { bubbles: true }));
    const minimum = fieldInput(yAxis, 'Minimum');
    minimum.value = '-5';
    minimum.dispatchEvent(new Event('input', { bubbles: true }));
    const maximum = fieldInput(yAxis, 'Maximum');
    maximum.value = '0';
    maximum.dispatchEvent(new Event('input', { bubbles: true }));
    const majorUnit = fieldInput(yAxis, 'Major unit');
    majorUnit.value = '0';
    majorUnit.dispatchEvent(new Event('input', { bubbles: true }));
    changedSurface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();

    const chart = store.getState().charts.charts[0];
    expect(chart?.xAxis).toMatchObject({ title: 'Updated X', customAxisValue: 'keep-x' });
    expect(chart?.yAxis).toMatchObject({
      min: -5,
      max: 0,
      numberFormat: '0.0',
      majorGridlines: false,
      customAxisValue: 'keep-y',
    });
    expect(chart?.yAxis?.majorUnit).toBeUndefined();
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.xAxis).toMatchObject({
      title: 'Original X',
      customAxisValue: 'keep-x',
    });
    handle.detach();
    host.remove();
  });

  it('applies both axis reverse-order controls sparsely with undo, redo, and cancel', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertChart(store, {
      id: 'axis-reverse-history',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      xAxis: { title: 'Original X', customAxisValue: 'keep-x' } as never,
      yAxis: { reverseOrder: false, title: 'Original Y', customAxisValue: 'keep-y' } as never,
    });
    const handle = attachSessionCharts({ host, store, history });
    const openElements = (): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };
    const reverseOrder = (surface: HTMLElement, axisLabel: string): HTMLInputElement => {
      const fieldset = Array.from(surface.querySelectorAll<HTMLElement>('fieldset')).find(
        (candidate) => candidate.querySelector('legend')?.textContent === axisLabel,
      );
      const row = Array.from(
        fieldset?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [],
      ).find((candidate) => candidate.textContent?.includes('Reverse order'));
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing reverse order checkbox: ${axisLabel}`);
      return input;
    };

    const before = {
      xAxis: store.getState().charts.charts[0]?.xAxis,
      yAxis: store.getState().charts.charts[0]?.yAxis,
    };
    openElements().querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis).toEqual(before.xAxis);
    expect(store.getState().charts.charts[0]?.yAxis).toEqual(before.yAxis);
    expect(history.canUndo()).toBe(false);

    const noOp = openElements();
    const xNoOp = reverseOrder(noOp, 'X-axis');
    const yNoOp = reverseOrder(noOp, 'Y-axis');
    xNoOp.checked = true;
    xNoOp.dispatchEvent(new Event('change', { bubbles: true }));
    xNoOp.checked = false;
    xNoOp.dispatchEvent(new Event('change', { bubbles: true }));
    yNoOp.checked = true;
    yNoOp.dispatchEvent(new Event('change', { bubbles: true }));
    yNoOp.checked = false;
    yNoOp.dispatchEvent(new Event('change', { bubbles: true }));
    noOp.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis).toEqual(before.xAxis);
    expect(history.canUndo()).toBe(false);

    const changed = openElements();
    const xChanged = reverseOrder(changed, 'X-axis');
    const yChanged = reverseOrder(changed, 'Y-axis');
    xChanged.checked = true;
    xChanged.dispatchEvent(new Event('change', { bubbles: true }));
    yChanged.checked = true;
    yChanged.dispatchEvent(new Event('change', { bubbles: true }));
    changed.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis).toMatchObject({
      title: 'Original X',
      customAxisValue: 'keep-x',
      reverseOrder: true,
    });
    expect(store.getState().charts.charts[0]?.yAxis).toMatchObject({
      title: 'Original Y',
      customAxisValue: 'keep-y',
      reverseOrder: true,
    });
    expect(history.canUndo()).toBe(true);

    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.xAxis).toEqual(before.xAxis);
    expect(store.getState().charts.charts[0]?.yAxis).toEqual(before.yAxis);
    expect(history.redo()).toBe(true);
    expect(store.getState().charts.charts[0]?.xAxis?.reverseOrder).toBe(true);
    expect(store.getState().charts.charts[0]?.yAxis?.reverseOrder).toBe(true);

    const cancel = openElements();
    const xCancel = reverseOrder(cancel, 'X-axis');
    xCancel.checked = false;
    xCancel.dispatchEvent(new Event('change', { bubbles: true }));
    cancel.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
    expect(store.getState().charts.charts[0]?.xAxis?.reverseOrder).toBe(true);
    expect(store.getState().charts.charts[0]?.yAxis?.reverseOrder).toBe(true);
    handle.detach();
    host.remove();
  });

  it('syncs untouched external axis reversal while preserving a dirty draft', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertChart(store, {
      id: 'axis-reverse-sync',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store, history });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing chart settings');
    const reverseOrder = (axisLabel: string): HTMLInputElement => {
      const fieldset = Array.from(surface.querySelectorAll<HTMLElement>('fieldset')).find(
        (candidate) => candidate.querySelector('legend')?.textContent === axisLabel,
      );
      const row = Array.from(
        fieldset?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [],
      ).find((candidate) => candidate.textContent?.includes('Reverse order'));
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing reverse order checkbox: ${axisLabel}`);
      return input;
    };

    const title = surface.querySelector<HTMLInputElement>(
      '.fc-chart-settings__field input[type="text"]',
    );
    if (!title) throw new Error('missing chart title input');
    title.value = 'Draft title';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    mutators.updateChart(store, 'axis-reverse-sync', { xAxis: { reverseOrder: true } });
    expect(reverseOrder('X-axis').checked).toBe(true);
    expect(
      surface.querySelector<HTMLInputElement>('.fc-chart-settings__field input[type="text"]')
        ?.value,
    ).toBe('Draft title');

    const yDraft = reverseOrder('Y-axis');
    yDraft.checked = true;
    yDraft.dispatchEvent(new Event('change', { bubbles: true }));
    mutators.updateChart(store, 'axis-reverse-sync', { yAxis: { reverseOrder: false } });
    expect(reverseOrder('Y-axis').checked).toBe(true);

    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.title).toBe('Draft title');
    expect(store.getState().charts.charts[0]?.xAxis?.reverseOrder).toBe(true);
    expect(store.getState().charts.charts[0]?.yAxis?.reverseOrder).toBe(true);
    handle.detach();
    host.remove();
  });

  it.each(['pie', 'doughnut'] as const)('disables reverse-order controls for %s charts', (kind) => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: `axis-reverse-${kind}`,
      kind,
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      xAxis: { reverseOrder: true },
      yAxis: { reverseOrder: false },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing chart settings');
    for (const axisLabel of ['X-axis', 'Y-axis']) {
      const fieldset = Array.from(surface.querySelectorAll<HTMLElement>('fieldset')).find(
        (candidate) => candidate.querySelector('legend')?.textContent === axisLabel,
      );
      const row = Array.from(
        fieldset?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [],
      ).find((candidate) => candidate.textContent?.includes('Reverse order'));
      const input = row?.querySelector<HTMLInputElement>('input');
      expect(input?.disabled).toBe(true);
    }
    handle.detach();
    host.remove();
  });

  it('disables axis controls according to chart roles and disables unsupported grouping', () => {
    const cases = [
      { kind: 'column' as const, xNumeric: false, yNumeric: true, grouping: false },
      { kind: 'bar' as const, xNumeric: true, yNumeric: false, grouping: false },
      { kind: 'scatter' as const, xNumeric: true, yNumeric: true, grouping: true },
      { kind: 'pie' as const, xNumeric: false, yNumeric: false, grouping: true },
    ];
    for (const item of cases) {
      const host = document.createElement('div');
      document.body.appendChild(host);
      const store = createSpreadsheetStore();
      mutators.upsertChart(store, {
        id: `axis-role-${item.kind}`,
        kind: item.kind,
        source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      });
      const handle = attachSessionCharts({ host, store });
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing settings surface');
      const axis = (label: string): HTMLElement => {
        const fieldset = Array.from(surface.querySelectorAll<HTMLElement>('fieldset')).find(
          (candidate) => candidate.querySelector('legend')?.textContent === label,
        );
        if (!fieldset) throw new Error(`missing axis ${label}`);
        return fieldset;
      };
      const minimum = (fieldset: HTMLElement): HTMLInputElement => {
        const field = Array.from(
          fieldset.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
        ).find(
          (candidate) =>
            !candidate.classList.contains('fc-chart-settings__crossing-mode') &&
            candidate.querySelector('.fc-chart-settings__label')?.textContent === 'Minimum',
        );
        const input = field?.querySelector<HTMLInputElement>('input');
        if (!input) throw new Error('missing minimum input');
        return input;
      };
      expect(minimum(axis('X-axis')).disabled).toBe(!item.xNumeric);
      expect(minimum(axis('Y-axis')).disabled).toBe(!item.yNumeric);
      const grouping = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((field) => field.textContent?.includes('Grouping'));
      expect(grouping?.querySelector('select')?.disabled).toBe(item.grouping);
      handle.detach();
      host.remove();
    }
  });

  it('shows automatic for invalid stored major units and keeps the raw value until edited', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'invalid-major-unit',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      yAxis: { majorUnit: -1, customAxisValue: 'keep' } as never,
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const yAxis = Array.from(surface?.querySelectorAll<HTMLElement>('fieldset') ?? []).find(
      (candidate) => candidate.querySelector('legend')?.textContent === 'Y-axis',
    );
    const majorUnit = Array.from(
      yAxis?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Major unit'))
      ?.querySelector<HTMLInputElement>('input');
    expect(majorUnit?.value).toBe('');
    const title = Array.from(
      yAxis?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Y-axis title'))
      ?.querySelector<HTMLInputElement>('input');
    if (!title) throw new Error('missing y-axis title');
    title.value = 'Value';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.yAxis).toMatchObject({
      majorUnit: -1,
      customAxisValue: 'keep',
      title: 'Value',
    });
    handle.detach();
    host.remove();
  });

  it('shows forced zero for stacked line empty cells without rewriting the stored choice', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'forced-empty-cells',
      kind: 'line',
      grouping: 'stacked',
      emptyCells: 'gap',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing settings surface');
    const field = Array.from(
      surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
    ).find((candidate) => candidate.textContent?.includes('Empty cells'));
    const select = field?.querySelector<HTMLSelectElement>('select');
    expect(select?.disabled).toBe(true);
    expect(select?.value).toBe('zero');
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.emptyCells).toBe('gap');
    handle.detach();
    host.remove();
  });

  it('enables scatter span only for line modes and preserves it through kind changes', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const table = [
      ['X', 'Sales'],
      [1, 10],
      [null, null],
      [3, 30],
    ];
    table.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        if (value === null) return;
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      {
        id: 'scatter-empty-cells',
        kind: 'scatter',
        seriesDirection: 'columns',
        emptyCells: 'gap',
      },
    );
    const handle = attachSessionCharts({ host, store, history });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing scatter elements');
    const selectFor = (label: string): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes(label));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error(`missing ${label} select`);
      return select;
    };
    const spanOption = (): HTMLOptionElement => {
      const option = Array.from(selectFor('Empty cells').options).find(
        (candidate) => candidate.value === 'span',
      );
      if (!option) throw new Error('missing span option');
      return option;
    };
    const chartType = (): HTMLSelectElement => selectFor('Chart type');
    const scatterStyle = (): HTMLSelectElement => selectFor('Scatter style');

    expect(scatterStyle().value).toBe('marker');
    expect(spanOption().disabled).toBe(true);
    scatterStyle().value = 'line';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    expect(spanOption().disabled).toBe(false);
    scatterStyle().value = 'lineMarker';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    expect(spanOption().disabled).toBe(false);

    scatterStyle().value = 'line';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    const emptyCells = selectFor('Empty cells');
    emptyCells.value = 'span';
    emptyCells.dispatchEvent(new Event('change', { bubbles: true }));
    chartType().value = 'column';
    chartType().dispatchEvent(new Event('change', { bubbles: true }));
    expect(selectFor('Empty cells').value).toBe('span');
    chartType().value = 'scatter';
    chartType().dispatchEvent(new Event('change', { bubbles: true }));
    expect(scatterStyle().value).toBe('line');
    expect(selectFor('Empty cells').value).toBe('span');
    expect(spanOption().disabled).toBe(false);

    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      kind: 'scatter',
      emptyCells: 'span',
      scatterStyle: 'line',
    });
    expect(host.querySelectorAll('[data-chart-role="line"]')).toHaveLength(1);
    handle.detach();
    host.remove();
  });

  it('applies chart geometry knobs with Excel ranges and keeps pie angle separate', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'knob-column',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing settings surface');
    const fieldInput = (text: string): HTMLInputElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes(text));
      const input = field?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing geometry field: ${text}`);
      return input;
    };
    const gap = fieldInput('Gap width');
    gap.value = '0';
    gap.dispatchEvent(new Event('input', { bubbles: true }));
    const overlap = fieldInput('Overlap');
    overlap.value = '-100';
    overlap.dispatchEvent(new Event('input', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({ gapWidth: 0, overlap: -100 });
    handle.detach();
    host.remove();

    const pieHost = document.createElement('div');
    document.body.appendChild(pieHost);
    const pieStore = createSpreadsheetStore();
    mutators.upsertChart(pieStore, {
      id: 'knob-pie',
      kind: 'pie',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const pieHandle = attachSessionCharts({ host: pieHost, store: pieStore });
    pieHost.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const pieSurface = pieHost.querySelector<HTMLElement>('.fc-chart-settings');
    const angle = Array.from(
      pieSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('First slice angle'))
      ?.querySelector<HTMLInputElement>('input');
    if (!angle) throw new Error('missing pie angle input');
    angle.value = '360';
    angle.dispatchEvent(new Event('input', { bubbles: true }));
    pieSurface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(pieStore.getState().charts.charts[0]?.firstSliceAngle).toBe(360);
    pieHandle.detach();
    pieHost.remove();
  });

  it('does not materialize clamped geometry defaults when another field changes', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'raw-geometry',
      kind: 'column',
      gapWidth: 999,
      overlap: 200,
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing chart title');
    title.value = 'Keep raw geometry';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      title: 'Keep raw geometry',
      gapWidth: 999,
      overlap: 200,
    });
    handle.detach();
    host.remove();
  });

  it('keeps an out-of-range doughnut hole raw until the field is edited', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'raw-hole-size',
      kind: 'doughnut',
      holeSize: 99,
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const elements = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = elements?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing title input');
    title.value = 'Keep raw hole';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    elements?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      title: 'Keep raw hole',
      holeSize: 99,
    });

    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const styles = host.querySelector<HTMLElement>('.fc-chart-settings');
    const hole = Array.from(
      styles?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Hole size'))
      ?.querySelector<HTMLInputElement>('input');
    if (!hole) throw new Error('missing hole size input');
    expect(hole.value).toBe('90');
    hole.value = '90';
    hole.dispatchEvent(new Event('input', { bubbles: true }));
    styles?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.holeSize).toBe(99);

    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const reopenedStyles = host.querySelector<HTMLElement>('.fc-chart-settings');
    const reopenedHole = Array.from(
      reopenedStyles?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Hole size'))
      ?.querySelector<HTMLInputElement>('input');
    if (!reopenedHole) throw new Error('missing reopened hole size input');
    reopenedHole.value = '70';
    reopenedHole.dispatchEvent(new Event('input', { bubbles: true }));
    reopenedStyles?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.holeSize).toBe(70);
    handle.detach();
    host.remove();

    const defaultHost = document.createElement('div');
    document.body.appendChild(defaultHost);
    const defaultStore = createSpreadsheetStore();
    const defaultHistory = new History();
    mutators.upsertChart(defaultStore, {
      id: 'default-hole-size',
      kind: 'doughnut',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const defaultHandle = attachSessionCharts({
      host: defaultHost,
      store: defaultStore,
      history: defaultHistory,
    });
    defaultHost.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const defaultStyles = defaultHost.querySelector<HTMLElement>('.fc-chart-settings');
    const defaultHole = Array.from(
      defaultStyles?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Hole size'))
      ?.querySelector<HTMLInputElement>('input');
    if (!defaultHole) throw new Error('missing default hole size input');
    defaultHole.value = '70';
    defaultHole.dispatchEvent(new Event('input', { bubbles: true }));
    defaultStyles?.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-elements')?.click();
    const defaultElements = defaultHost.querySelector<HTMLElement>('.fc-chart-settings');
    defaultElements?.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    const rerenderedStyles = defaultHost.querySelector<HTMLElement>('.fc-chart-settings');
    const rerenderedHole = Array.from(
      rerenderedStyles?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Hole size'))
      ?.querySelector<HTMLInputElement>('input');
    if (!rerenderedHole) throw new Error('missing rerendered hole size input');
    rerenderedHole.value = '50';
    rerenderedHole.dispatchEvent(new Event('input', { bubbles: true }));
    rerenderedStyles?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(defaultStore.getState().charts.charts[0]?.holeSize).toBeUndefined();
    expect(defaultHistory.canUndo()).toBe(false);
    defaultHandle.detach();
    defaultHost.remove();
  });

  it('reparses draft data across chart kinds and applies one undoable kind change', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const table = [
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 6],
      ['Feb', 15, 8],
      ['Mar', 8, 3],
    ];
    table.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 },
      { id: 'draft-kind-data', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store, history });
    const chartBefore = structuredClone(store.getState().charts.charts[0]);
    const openElements = (): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };
    const chartType = (surface: HTMLElement): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes('Chart type'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing chart type select');
      return select;
    };

    const cancelled = openElements();
    chartType(cancelled).value = 'scatter';
    chartType(cancelled).dispatchEvent(new Event('change', { bubbles: true }));
    cancelled.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
    expect(store.getState().charts.charts[0]).toEqual(chartBefore);

    const scatterDraft = openElements();
    chartType(scatterDraft).value = 'scatter';
    chartType(scatterDraft).dispatchEvent(new Event('change', { bubbles: true }));
    scatterDraft.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    let styles = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!styles) throw new Error('missing scatter styles');
    const scatterSeries = Array.from(
      styles.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
    )
      .find((field) => field.querySelector('.fc-chart-settings__label')?.textContent === 'Series')
      ?.querySelector<HTMLSelectElement>('select');
    const scatterPoints = Array.from(
      styles.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
    )
      .find((field) => field.querySelector('.fc-chart-settings__label')?.textContent === 'Point')
      ?.querySelector<HTMLSelectElement>('select');
    expect(Array.from(scatterSeries?.options ?? [])).toHaveLength(1);
    expect(scatterSeries?.options[0]?.textContent).toBe('Costs');
    expect(Array.from(scatterPoints?.options ?? []).map((option) => option.value)).toEqual([
      '__all__',
      'point-0',
      'point-1',
      'point-2',
    ]);

    styles.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-elements')?.click();
    const doughnutDraft = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!doughnutDraft) throw new Error('missing doughnut elements');
    chartType(doughnutDraft).value = 'doughnut';
    chartType(doughnutDraft).dispatchEvent(new Event('change', { bubbles: true }));
    doughnutDraft.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    styles = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!styles) throw new Error('missing doughnut styles');
    const doughnutSeries = Array.from(
      styles.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
    )
      .find((field) => field.querySelector('.fc-chart-settings__label')?.textContent === 'Series')
      ?.querySelector<HTMLSelectElement>('select');
    const doughnutPoints = Array.from(
      styles.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
    )
      .find((field) => field.querySelector('.fc-chart-settings__label')?.textContent === 'Point')
      ?.querySelector<HTMLSelectElement>('select');
    expect(Array.from(doughnutSeries?.options ?? [])).toHaveLength(2);
    expect(Array.from(doughnutPoints?.options ?? []).map((option) => option.value)).toEqual([
      '__all__',
      'category-0',
      'category-1',
      'category-2',
    ]);
    styles.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.kind).toBe('doughnut');
    expect(host.querySelectorAll('[data-chart-role="slice"]')).toHaveLength(6);
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.kind).toBe('column');
    expect(history.redo()).toBe(true);
    expect(store.getState().charts.charts[0]?.kind).toBe('doughnut');
    handle.detach();
    host.remove();
  });

  it('keeps implicit legend and vary-colors defaults sparse across draft toggles', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const table = [
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 6],
      ['Feb', 15, 8],
    ];
    table.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      { id: 'implicit-chart-defaults', kind: 'pie', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store, history });
    const checkbox = (surface: HTMLElement, label: string): HTMLInputElement => {
      const row = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__check'),
      ).find((candidate) => candidate.textContent?.trim() === label);
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing checkbox: ${label}`);
      return input;
    };
    const open = (tab: 'elements' | 'styles'): HTMLElement => {
      host.querySelector<HTMLButtonElement>(`[data-chart-settings-tab="${tab}"]`)?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };

    let surface = open('elements');
    const legend = checkbox(surface, 'Legend');
    expect(legend.checked).toBe(true);
    legend.checked = false;
    legend.dispatchEvent(new Event('change', { bubbles: true }));
    legend.checked = true;
    legend.dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.legend).toBeUndefined();
    expect(history.canUndo()).toBe(false);

    surface = open('styles');
    const vary = checkbox(surface, 'Vary colors by point');
    expect(vary.checked).toBe(true);
    vary.checked = false;
    vary.dispatchEvent(new Event('change', { bubbles: true }));
    vary.checked = true;
    vary.dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.varyColors).toBeUndefined();
    expect(history.canUndo()).toBe(false);

    surface = open('elements');
    const kindField = Array.from(
      surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
    ).find((field) => field.textContent?.includes('Chart type'));
    const kind = kindField?.querySelector<HTMLSelectElement>('select');
    if (!kind) throw new Error('missing chart type select');
    kind.value = 'doughnut';
    kind.dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    const explicitFalse = checkbox(surface, 'Vary colors by point');
    expect(explicitFalse.checked).toBe(true);
    explicitFalse.checked = false;
    explicitFalse.dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      kind: 'doughnut',
      varyColors: false,
    });
    expect(history.canUndo()).toBe(true);
    handle.detach();
    host.remove();
  });

  it('keeps optional defaults sparse when a new kind checkbox is toggled back', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    [
      ['Month', 'Sales'],
      ['Jan', 10],
      ['Feb', 20],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      { id: 'sparse-kind-defaults', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store, history });
    const checkbox = (surface: HTMLElement, label: string): HTMLInputElement => {
      const row = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__check'),
      ).find((candidate) => candidate.textContent?.trim() === label);
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing checkbox: ${label}`);
      return input;
    };
    const chartType = (surface: HTMLElement): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes('Chart type'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing chart type select');
      return select;
    };
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    let surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing chart settings');
    expect(checkbox(surface, 'Legend').checked).toBe(false);
    chartType(surface).value = 'doughnut';
    chartType(surface).dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing chart settings after kind change');

    const legend = checkbox(surface, 'Legend');
    expect(legend.checked).toBe(true);
    legend.checked = false;
    legend.dispatchEvent(new Event('change', { bubbles: true }));
    legend.checked = true;
    legend.dispatchEvent(new Event('change', { bubbles: true }));

    surface.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing styles after kind change');
    const vary = checkbox(surface, 'Vary colors by point');
    expect(vary.checked).toBe(true);
    vary.checked = false;
    vary.dispatchEvent(new Event('change', { bubbles: true }));
    vary.checked = true;
    vary.dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-elements')?.click();
    surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing elements after kind change');
    chartType(surface).value = 'column';
    chartType(surface).dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({ kind: 'column' });
    expect(store.getState().charts.charts[0]?.legend).toBeUndefined();
    expect(store.getState().charts.charts[0]?.varyColors).toBeUndefined();
    expect(history.canUndo()).toBe(false);
    handle.detach();
    host.remove();

    const explicitHost = document.createElement('div');
    document.body.appendChild(explicitHost);
    const explicitStore = createSpreadsheetStore();
    const explicitHistory = new History();
    mutators.setCell(explicitStore, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'Month' });
    mutators.setCell(explicitStore, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'Sales' });
    mutators.setCell(explicitStore, { sheet: 0, row: 1, col: 0 }, { kind: 'text', value: 'Jan' });
    mutators.setCell(explicitStore, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 10 });
    createSessionChart(
      explicitStore,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'explicit-vary-colors', kind: 'column', varyColors: false },
    );
    const explicitHandle = attachSessionCharts({
      host: explicitHost,
      store: explicitStore,
      history: explicitHistory,
    });
    explicitHost.querySelector<HTMLButtonElement>('[data-chart-settings-tab="styles"]')?.click();
    const explicitSurface = explicitHost.querySelector<HTMLElement>('.fc-chart-settings');
    if (!explicitSurface) throw new Error('missing explicit styles');
    const explicitVary = checkbox(explicitSurface, 'Vary colors by point');
    expect(explicitVary.checked).toBe(false);
    explicitVary.checked = true;
    explicitVary.dispatchEvent(new Event('change', { bubbles: true }));
    explicitVary.checked = false;
    explicitVary.dispatchEvent(new Event('change', { bubbles: true }));
    explicitSurface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(explicitStore.getState().charts.charts[0]?.varyColors).toBe(false);
    expect(explicitHistory.canUndo()).toBe(false);
    explicitHandle.detach();
    explicitHost.remove();
  });

  it('does not overwrite an external kind update when applying another draft field', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'external-kind-update',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing title input');
    title.value = 'External kind survives';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    mutators.updateChart(store, 'external-kind-update', { kind: 'scatter' });
    surface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      kind: 'scatter',
      title: 'External kind survives',
    });
    handle.detach();
    host.remove();
  });

  it('keeps external legend and vary-colors overrides when applying only a title', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertChart(store, {
      id: 'external-optional-updates',
      kind: 'column',
      legend: { visible: false },
      varyColors: false,
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store, history });
    const checkbox = (surface: HTMLElement, label: string): HTMLInputElement => {
      const row = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__check'),
      ).find((candidate) => candidate.textContent?.trim() === label);
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing checkbox: ${label}`);
      return input;
    };
    const openAndEditTitle = (value: string): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      const title = surface?.querySelector<HTMLInputElement>('input[type="text"]');
      if (!surface || !title) throw new Error('missing title settings');
      title.value = value;
      title.dispatchEvent(new Event('input', { bubbles: true }));
      return surface;
    };

    let surface = openAndEditTitle('External true');
    mutators.updateChart(store, 'external-optional-updates', {
      legend: { visible: true },
      varyColors: true,
    });
    expect(checkbox(surface, 'Legend').checked).toBe(true);
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      title: 'External true',
      legend: { visible: true },
      varyColors: true,
    });

    surface = openAndEditTitle('External false');
    surface.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    expect(checkbox(surface, 'Vary colors by point').checked).toBe(true);
    mutators.updateChart(store, 'external-optional-updates', {
      legend: { visible: false },
      varyColors: false,
    });
    expect(checkbox(surface, 'Vary colors by point').checked).toBe(false);
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      title: 'External false',
      legend: { visible: false },
      varyColors: false,
    });
    handle.detach();
    host.remove();
  });

  it('refreshes chart kind and axis roles for external same-domain kind changes', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    [
      ['Month', 'Sales'],
      ['Jan', 10],
      ['Feb', 20],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      { id: 'external-same-domain-kind', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();

    const chartType = (): HTMLSelectElement => {
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      const field = Array.from(
        surface?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
      ).find((candidate) => candidate.textContent?.includes('Chart type'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing chart type select');
      return select;
    };
    const axisMinimum = (axis: 'X-axis' | 'Y-axis'): HTMLInputElement => {
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      const fieldset = Array.from(surface?.querySelectorAll<HTMLElement>('fieldset') ?? []).find(
        (candidate) => candidate.querySelector('legend')?.textContent === axis,
      );
      const field = Array.from(
        fieldset?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
      ).find(
        (candidate) =>
          !candidate.classList.contains('fc-chart-settings__crossing-mode') &&
          candidate.querySelector('.fc-chart-settings__label')?.textContent === 'Minimum',
      );
      const input = field?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error(`missing ${axis} minimum input`);
      return input;
    };

    expect(chartType().value).toBe('column');
    expect(axisMinimum('X-axis').disabled).toBe(true);
    expect(axisMinimum('Y-axis').disabled).toBe(false);

    mutators.updateChart(store, 'external-same-domain-kind', { kind: 'bar' });
    expect(chartType().value).toBe('bar');
    expect(axisMinimum('X-axis').disabled).toBe(false);
    expect(axisMinimum('Y-axis').disabled).toBe(true);

    mutators.updateChart(store, 'external-same-domain-kind', { kind: 'line' });
    expect(chartType().value).toBe('line');
    expect(axisMinimum('X-axis').disabled).toBe(true);
    expect(axisMinimum('Y-axis').disabled).toBe(false);
    handle.detach();
    host.remove();
  });

  it('shows all scatter styles and resolves marker visibility for each mode', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    [
      ['X', 'Sales'],
      [1, 10],
      [2, 20],
      [3, 30],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      { id: 'scatter-style-modes', kind: 'scatter', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing scatter elements');
    const scatterStyle = (): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes('Scatter style'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing scatter style select');
      return select;
    };
    expect(Array.from(scatterStyle().options).map((option) => option.value)).toEqual([
      'marker',
      'line',
      'lineMarker',
    ]);
    expect(scatterStyle().value).toBe('marker');

    scatterStyle().value = 'line';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    const lineSurface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const lineMarker = Array.from(
      lineSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [],
    ).find((candidate) => candidate.textContent?.includes('Visible'));
    const lineMarkerInput = lineMarker?.querySelector<HTMLInputElement>('input');
    expect(lineMarkerInput?.checked).toBe(false);
    if (!lineMarkerInput) throw new Error('missing line marker checkbox');
    lineMarkerInput.checked = true;
    lineMarkerInput.dispatchEvent(new Event('change', { bubbles: true }));

    lineSurface?.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-elements')?.click();
    scatterStyle().value = 'lineMarker';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('#fc-chart-settings-tab-styles')?.click();
    const lineMarkerSurface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const lineMarkerCheckbox = Array.from(
      lineMarkerSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [],
    ).find((candidate) => candidate.textContent?.includes('Visible'));
    expect(lineMarkerCheckbox?.querySelector<HTMLInputElement>('input')?.checked).toBe(true);
    lineMarkerSurface?.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.seriesStyles?.[0]?.marker?.visible).toBe(true);
    handle.detach();
    host.remove();
  });

  it('keeps scatter style sparse through kind cycles and preserves dormant edits', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'X' });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'Sales' });
    mutators.setCell(store, { sheet: 0, row: 1, col: 0 }, { kind: 'number', value: 1 });
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 10 });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'scatter-style-dormant', kind: 'column', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store, history });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing kind settings');
    const chartType = (): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes('Chart type'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing chart type select');
      return select;
    };
    const scatterStyle = (): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes('Scatter style'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing scatter style select');
      return select;
    };
    chartType().value = 'scatter';
    chartType().dispatchEvent(new Event('change', { bubbles: true }));
    expect(scatterStyle().value).toBe('marker');
    scatterStyle().value = 'line';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    chartType().value = 'column';
    chartType().dispatchEvent(new Event('change', { bubbles: true }));
    chartType().value = 'scatter';
    chartType().dispatchEvent(new Event('change', { bubbles: true }));
    expect(scatterStyle().value).toBe('line');
    scatterStyle().value = 'marker';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    chartType().value = 'column';
    chartType().dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({ kind: 'column' });
    expect(store.getState().charts.charts[0]?.scatterStyle).toBeUndefined();
    expect(history.canUndo()).toBe(false);
    handle.detach();
    host.remove();
  });

  it('preserves external scatter styles, no-op drafts, and Apply undo/redo', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'X' });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'Sales' });
    mutators.setCell(store, { sheet: 0, row: 1, col: 0 }, { kind: 'number', value: 1 });
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 10 });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      { id: 'scatter-style-history', kind: 'scatter', seriesDirection: 'columns' },
    );
    const handle = attachSessionCharts({ host, store, history });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing scatter elements');
    const scatterStyle = (): HTMLSelectElement => {
      const field = Array.from(
        surface.querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      ).find((candidate) => candidate.textContent?.includes('Scatter style'));
      const select = field?.querySelector<HTMLSelectElement>('select');
      if (!select) throw new Error('missing scatter style select');
      return select;
    };
    scatterStyle().value = 'line';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    scatterStyle().value = 'marker';
    scatterStyle().dispatchEvent(new Event('change', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.scatterStyle).toBeUndefined();
    expect(history.canUndo()).toBe(false);

    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    mutators.updateChart(store, 'scatter-style-history', { scatterStyle: 'line' });
    expect(scatterStyle().value).toBe('line');
    const title = surface.querySelector<HTMLInputElement>('input[type="text"]');
    if (!title) throw new Error('missing scatter title');
    title.value = 'External line';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    mutators.updateChart(store, 'scatter-style-history', { scatterStyle: 'lineMarker' });
    expect(scatterStyle().value).toBe('lineMarker');
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]).toMatchObject({
      title: 'External line',
      scatterStyle: 'lineMarker',
    });
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.scatterStyle).toBe('lineMarker');
    expect(history.redo()).toBe(true);
    expect(store.getState().charts.charts[0]?.title).toBe('External line');

    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const cancelSurface = host.querySelector<HTMLElement>('.fc-chart-settings');
    const cancelStyle = Array.from(
      cancelSurface?.querySelectorAll<HTMLElement>('.fc-chart-settings__field') ?? [],
    )
      .find((field) => field.textContent?.includes('Scatter style'))
      ?.querySelector<HTMLSelectElement>('select');
    if (cancelStyle) {
      cancelStyle.value = 'marker';
      cancelStyle.dispatchEvent(new Event('change', { bubbles: true }));
    }
    cancelSurface?.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
    expect(store.getState().charts.charts[0]?.scatterStyle).toBe('lineMarker');
    handle.detach();
    host.remove();
  });

  it.each(['column', 'bar', 'line', 'area', 'pie', 'scatter'] as const)(
    'renders %s through the chart scene and shows the title once',
    (kind) => {
      const host = document.createElement('div');
      const store = createSpreadsheetStore();
      const table =
        kind === 'scatter'
          ? [
              ['X', 'Sales'],
              [1, 10],
              [10, 20],
              [100, 30],
            ]
          : [
              ['Month', 'Sales'],
              ['Jan', 10],
              ['Feb', 20],
              ['Mar', 30],
            ];
      table.forEach((row, r) => {
        row.forEach((value, col) => {
          mutators.setCell(
            store,
            { sheet: 0, row: r, col },
            typeof value === 'number'
              ? { kind: 'number', value }
              : { kind: 'text', value: String(value) },
          );
        });
      });
      createSessionChart(
        store,
        { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
        {
          kind,
          title: 'Revenue',
          seriesDirection: 'columns',
        },
      );
      const handle = attachSessionCharts({ host, store });
      const role = kind === 'pie' ? 'slice' : kind === 'bar' || kind === 'column' ? 'bar' : 'point';
      expect(host.querySelectorAll(`[data-chart-role="${role}"]`)).toHaveLength(3);
      expect(
        host.querySelector(`[data-chart-role="${role}"]`)?.getAttribute('data-series-index'),
      ).toBe('0');
      expect(host.querySelector('svg')?.textContent).not.toContain('Revenue');
      expect(host.querySelector('.fc-chart__title')?.textContent).toBe('Revenue');
      handle.detach();
    },
  );

  it('preserves distinct series and updates chart marks when source cells change', () => {
    const host = document.createElement('div');
    const store = createSpreadsheetStore();
    const table = [
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
      ['Feb', 30, 40],
    ];
    table.forEach((row, r) => {
      row.forEach((value, col) => {
        mutators.setCell(
          store,
          { sheet: 0, row: r, col },
          typeof value === 'number'
            ? { kind: 'number', value }
            : { kind: 'text', value: String(value) },
        );
      });
    });
    createSessionChart(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
      {
        kind: 'column',
        seriesDirection: 'columns',
        yAxis: { min: 0, max: 100 },
      },
    );
    const handle = attachSessionCharts({ host, store });
    expect(host.querySelectorAll('[data-chart-role="bar"]')).toHaveLength(4);
    expect(host.querySelectorAll('text[data-chart-role="legend"]')).toHaveLength(2);
    const selector = '[data-chart-role="bar"][data-series-index="0"][data-point-index="0"]';
    const before = Number(host.querySelector(selector)?.getAttribute('height'));
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 50 });
    const after = Number(host.querySelector(selector)?.getAttribute('height'));
    expect(before).toBeGreaterThan(0);
    expect(after).toBeCloseTo(before * 5);
    handle.detach();
  });

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

  it('keeps sparse crossing auto and invalid raw values untouched until edited', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertChart(store, {
      id: 'crossing-sparse',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store, history });
    const openElements = (): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };

    let surface = openElements();
    expect(chartAxisCrossingSelect(surface, 'X-axis').value).toBe('auto');
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(history.canUndo()).toBe(false);
    expect(store.getState().charts.charts[0]?.xAxis).toBeUndefined();

    surface = openElements();
    const xSelect = chartAxisCrossingSelect(surface, 'X-axis');
    xSelect.value = 'min';
    xSelect.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    const xAuto = chartAxisCrossingSelect(surface, 'X-axis');
    xAuto.value = 'auto';
    xAuto.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(history.canUndo()).toBe(false);
    expect(store.getState().charts.charts[0]?.xAxis).toBeUndefined();

    mutators.updateChart(store, 'crossing-sparse', {
      xAxis: { crossing: Number.NaN } as never,
      yAxis: { crossing: Number.POSITIVE_INFINITY } as never,
    });
    surface = openElements();
    expect(chartAxisCrossingSelect(surface, 'X-axis').value).toBe('auto');
    expect(chartAxisCrossingSelect(surface, 'Y-axis').value).toBe('auto');
    const rawTitle = surface.querySelector<HTMLInputElement>(
      '.fc-chart-settings__field input[type="text"]',
    );
    if (!rawTitle) throw new Error('missing title input');
    rawTitle.value = 'Raw title';
    rawTitle.dispatchEvent(new Event('input', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    const untouched = store.getState().charts.charts[0];
    expect(Number.isNaN(untouched?.xAxis?.crossing as number)).toBe(true);
    expect(untouched?.yAxis?.crossing).toBe(Number.POSITIVE_INFINITY);
    expect(untouched?.title).toBe('Raw title');
    expect(history.canUndo()).toBe(true);
    history.clear();

    mutators.updateChart(store, 'crossing-sparse', {
      xAxis: { crossing: 'auto' },
      yAxis: { crossing: 'auto' },
    });
    surface = openElements();
    const explicitAuto = chartAxisCrossingSelect(surface, 'X-axis');
    explicitAuto.value = 'min';
    explicitAuto.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    const restoredAuto = chartAxisCrossingSelect(surface, 'X-axis');
    restoredAuto.value = 'auto';
    restoredAuto.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(history.canUndo()).toBe(false);
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe('auto');

    handle.detach();
    host.remove();
  });

  it('normalizes category crossings, keeps numeric defaults, and supports history and cancel', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    const table = [
      ['Category', 'Value'],
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ];
    table.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        mutators.setCell(
          store,
          { sheet: 0, row: rowIndex, col: colIndex },
          typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        );
      });
    });
    mutators.upsertChart(store, {
      id: 'crossing-values',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
      filter: { hiddenCategoryIds: ['category-1'] },
    });
    const handle = attachSessionCharts({ host, store, history });
    const openElements = (): HTMLElement => {
      host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
      const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
      if (!surface) throw new Error('missing chart settings');
      return surface;
    };

    let surface = openElements();
    const xMode = chartAxisCrossingSelect(surface, 'X-axis');
    xMode.value = 'custom';
    xMode.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    const xCustom = chartAxisCrossingCustom(surface, 'X-axis');
    expect(xCustom.min).toBe('1');
    expect(xCustom.step).toBe('1');
    xCustom.value = '2.6';
    xCustom.dispatchEvent(new Event('input', { bubbles: true }));

    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    const yMode = chartAxisCrossingSelect(surface, 'Y-axis');
    yMode.value = 'custom';
    yMode.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    expect(chartAxisCrossingCustom(surface, 'X-axis').value).toBe('2');
    const yCustom = chartAxisCrossingCustom(surface, 'Y-axis');
    expect(yCustom.step).toBe('any');
    expect(yCustom.value).toBe('0');
    yCustom.value = '-2.5';
    yCustom.dispatchEvent(new Event('input', { bubbles: true }));
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();

    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe(2);
    expect(store.getState().charts.charts[0]?.yAxis?.crossing).toBe(-2.5);
    expect(history.canUndo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBeUndefined();
    expect(store.getState().charts.charts[0]?.yAxis?.crossing).toBeUndefined();
    expect(history.redo()).toBe(true);
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe(2);
    expect(store.getState().charts.charts[0]?.yAxis?.crossing).toBe(-2.5);

    surface = openElements();
    const rememberedMin = chartAxisCrossingSelect(surface, 'X-axis');
    rememberedMin.value = 'min';
    rememberedMin.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    const rememberedCustom = chartAxisCrossingSelect(surface, 'X-axis');
    rememberedCustom.value = 'custom';
    rememberedCustom.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    expect(chartAxisCrossingCustom(surface, 'X-axis').value).toBe('2');
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();

    surface = openElements();
    const finiteMin = chartAxisCrossingSelect(surface, 'X-axis');
    finiteMin.value = 'min';
    finiteMin.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe('min');

    surface = openElements();
    const finiteMax = chartAxisCrossingSelect(surface, 'X-axis');
    finiteMax.value = 'max';
    finiteMax.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe('max');

    surface = openElements();
    const finiteAuto = chartAxisCrossingSelect(surface, 'X-axis');
    finiteAuto.value = 'auto';
    finiteAuto.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBeUndefined();

    surface = openElements();
    const cancelX = chartAxisCrossingSelect(surface, 'X-axis');
    cancelX.value = 'min';
    cancelX.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBeUndefined();

    handle.detach();
    host.remove();
  });

  it('retains dormant circular crossings and maps numeric roles per chart family', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertChart(store, {
      id: 'crossing-dormant',
      kind: 'pie',
      source: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
      xAxis: { crossing: 2 },
      yAxis: { crossing: 'max' },
    });
    const handle = attachSessionCharts({ host, store });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    const surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing chart settings');
    expect(chartAxisCrossingSelect(surface, 'X-axis').disabled).toBe(true);
    expect(chartAxisCrossingSelect(surface, 'Y-axis').disabled).toBe(true);
    expect(chartAxisCrossingCustom(surface, 'X-axis').disabled).toBe(true);
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe(2);
    expect(store.getState().charts.charts[0]?.yAxis?.crossing).toBe('max');
    handle.detach();
    host.remove();
  });

  it('syncs untouched crossings and preserves a dirty crossing with other draft fields', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'Category' });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'text', value: 'Value' });
    mutators.setCell(store, { sheet: 0, row: 1, col: 0 }, { kind: 'text', value: 'A' });
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'number', value: 1 });
    mutators.setCell(store, { sheet: 0, row: 2, col: 0 }, { kind: 'text', value: 'B' });
    mutators.setCell(store, { sheet: 0, row: 2, col: 1 }, { kind: 'number', value: 2 });
    mutators.upsertChart(store, {
      id: 'crossing-sync',
      kind: 'column',
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
    });
    const handle = attachSessionCharts({ host, store, history });
    host.querySelector<HTMLButtonElement>('[data-chart-settings-tab="elements"]')?.click();
    let surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing chart settings');
    expect(chartAxisCrossingSelect(surface, 'X-axis').value).toBe('auto');
    mutators.updateChart(store, 'crossing-sync', { xAxis: { crossing: 'min' } });
    surface = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!surface) throw new Error('missing synced chart settings');
    expect(chartAxisCrossingSelect(surface, 'X-axis').value).toBe('min');

    const title = surface.querySelector<HTMLInputElement>(
      '.fc-chart-settings__field input[type="text"]',
    );
    if (!title) throw new Error('missing title input');
    title.value = 'Draft title';
    title.dispatchEvent(new Event('input', { bubbles: true }));
    const dirtyMode = chartAxisCrossingSelect(surface, 'X-axis');
    dirtyMode.value = 'custom';
    dirtyMode.dispatchEvent(new Event('change', { bubbles: true }));
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    const dirtyCustom = chartAxisCrossingCustom(surface, 'X-axis');
    dirtyCustom.value = '2';
    dirtyCustom.dispatchEvent(new Event('input', { bubbles: true }));
    mutators.updateChart(store, 'crossing-sync', { xAxis: { crossing: 'max' } });
    surface = host.querySelector<HTMLElement>('.fc-chart-settings') as HTMLElement;
    expect(chartAxisCrossingSelect(surface, 'X-axis').value).toBe('custom');
    expect(
      surface.querySelector<HTMLInputElement>('.fc-chart-settings__field input[type="text"]')
        ?.value,
    ).toBe('Draft title');
    surface.querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
    expect(store.getState().charts.charts[0]?.title).toBe('Draft title');
    expect(store.getState().charts.charts[0]?.xAxis?.crossing).toBe(2);

    handle.detach();
    host.remove();
  });
});
