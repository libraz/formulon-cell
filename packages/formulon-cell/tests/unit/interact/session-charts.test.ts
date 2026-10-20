import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { AnchorEditAs, AnchorKind } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { automaticPalette, DEFAULT_THEME_ACCENTS } from '../../../src/charts/series-colors.js';
import type { ChartData, ChartKind } from '../../../src/charts/types.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { viewerPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  type CreateSessionChartOptions,
  createSessionChart,
  type SessionChartPatch,
  updateSessionChart,
} from '../../../src/commands/session-chart.js';
import {
  hydrateChartsFromEngine,
  refreshChartsAfterRecalc,
} from '../../../src/engine/charts-sync.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { formatSessionChartValue } from '../../../src/interact/session-chart-format.js';
import { attachSessionCharts } from '../../../src/interact/session-charts.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionChart,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const projection = vi.hoisted(() => ({
  /** Id of a chart seeded straight into the store, with no workbook object behind it. */
  id: 'chart-native-0-999',
  /** Every patch the settings surface applied, in order. */
  patches: [] as unknown[],
}));

// The settings surface's own contract is the patch it applies. A chart seeded
// only into the store stands in for the engine: its patch is echoed into the
// store so that sync and reopen behave as after a real update.
vi.mock('../../../src/commands/session-chart.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../../src/commands/session-chart.js')>();
  const { mutators: storeMutators } = await import('../../../src/store/store.js');
  return {
    ...actual,
    updateSessionChart: (...args: Parameters<typeof actual.updateSessionChart>) => {
      const [store, id, patch] = args;
      projection.patches.push(patch);
      if (id !== projection.id) return actual.updateSessionChart(...args);
      storeMutators.updateChart(store, id, patch as Partial<SessionChart>);
      return null;
    },
  };
});

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

const chartAxisLabelPositionSelect = (
  surface: HTMLElement,
  axisLabel: string,
): HTMLSelectElement => {
  const input = chartAxisFieldset(surface, axisLabel).querySelector<HTMLSelectElement>(
    '.fc-chart-settings__label-position select',
  );
  if (!input) throw new Error(`missing label position select: ${axisLabel}`);
  return input;
};

const chartAxisLogCheckbox = (surface: HTMLElement, axisLabel: string): HTMLInputElement => {
  const input = chartAxisFieldset(surface, axisLabel).querySelector<HTMLInputElement>(
    '.fc-chart-settings__log-scale input',
  );
  if (!input) throw new Error(`missing logarithmic scale checkbox: ${axisLabel}`);
  return input;
};

const chartAxisLogBase = (surface: HTMLElement, axisLabel: string): HTMLInputElement => {
  const input = chartAxisFieldset(surface, axisLabel).querySelector<HTMLInputElement>(
    '.fc-chart-settings__log-base input',
  );
  if (!input) throw new Error(`missing log base input: ${axisLabel}`);
  return input;
};

const setInput = (input: HTMLInputElement, value: string): void => {
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
};

const setSelect = (select: HTMLSelectElement, value: string): void => {
  select.value = value;
  select.dispatchEvent(new Event('change', { bubbles: true }));
};

const setChecked = (input: HTMLInputElement, checked: boolean): void => {
  input.checked = checked;
  input.dispatchEvent(new Event('change', { bubbles: true }));
};

/** A field control whose label row contains `label`. */
const fieldControl = <T extends HTMLElement = HTMLInputElement>(
  within: HTMLElement,
  label: string,
  selector = 'input',
): T => {
  const field = Array.from(within.querySelectorAll<HTMLElement>('.fc-chart-settings__field')).find(
    (candidate) => candidate.textContent?.includes(label),
  );
  const control = field?.querySelector<T>(selector);
  if (!control) throw new Error(`missing field: ${label}`);
  return control;
};

/** A checkbox whose label is exactly `label`. */
const checkbox = (within: HTMLElement, label: string): HTMLInputElement => {
  const row = Array.from(within.querySelectorAll<HTMLElement>('.fc-chart-settings__check')).find(
    (candidate) => candidate.textContent?.trim() === label,
  );
  const input = row?.querySelector<HTMLInputElement>('input');
  if (!input) throw new Error(`missing checkbox: ${label}`);
  return input;
};

const reverseOrder = (surface: HTMLElement, axisLabel: string): HTMLInputElement => {
  const row = Array.from(
    chartAxisFieldset(surface, axisLabel).querySelectorAll<HTMLElement>(
      '.fc-chart-settings__check',
    ),
  ).find((candidate) => candidate.textContent?.includes('Reverse order'));
  const input = row?.querySelector<HTMLInputElement>('input');
  if (!input) throw new Error(`missing reverse order checkbox: ${axisLabel}`);
  return input;
};

const CATEGORY_DATA: ChartData = {
  domain: 'category',
  categories: ['A', 'B'],
  categoryIds: ['category-0', 'category-1'],
  series: [{ id: 'series-0', name: 'Value', values: [10, 20] }],
};

const PROJECTED_ANCHOR = {
  anchorKind: AnchorKind.TwoCell,
  editAs: AnchorEditAs.TwoCell,
  fromRow: 1,
  fromCol: 1,
  fromRowOff: 0,
  fromColOff: 0,
  toRow: 16,
  toCol: 8,
  toRowOff: 0,
  toColOff: 0,
  cx: 4_572_000,
  cy: 2_743_200,
};

describe('attachSessionCharts', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;
  let host: HTMLDivElement;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    store = createSpreadsheetStore();
    history = new History();
    host = document.createElement('div');
    document.body.appendChild(host);
    projection.patches.length = 0;
  });

  afterEach(() => {
    host.remove();
    wb.dispose();
  });

  const range = (r0: number, c0: number, r1: number, c1: number): Range => ({
    sheet: 0,
    r0,
    c0,
    r1,
    c1,
  });

  /** Write a table into the workbook from A1. */
  const put = (rows: readonly (readonly (string | number | null)[])[]): void => {
    rows.forEach((cells, row) => {
      cells.forEach((value, col) => {
        if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col }, value);
        else if (typeof value === 'string') wb.setText({ sheet: 0, row, col }, value);
      });
    });
  };

  const create = (
    source: Range,
    options: Omit<CreateSessionChartOptions, 'wb'> = {},
  ): SessionChart => {
    const chart = createSessionChart(store, source, { wb, ...options });
    if (!chart) throw new Error('chart was not created');
    return chart;
  };

  /** Seed a chart into the store only; see the module mock. */
  const project = (
    spec: Partial<Omit<SessionChart, 'kind'>> & { kind: ChartKind },
    data: ChartData = CATEGORY_DATA,
  ): void => {
    mutators.upsertChart(store, {
      id: projection.id,
      sheet: 0,
      nativeObjectId: 999,
      nativeAnchor: PROJECTED_ANCHOR,
      data,
      ...spec,
    });
  };

  const attach = (
    extra: Partial<Parameters<typeof attachSessionCharts>[0]> = {},
  ): ReturnType<typeof attachSessionCharts> =>
    attachSessionCharts({ host, store, history, getWb: () => wb, ...extra });

  const surface = (): HTMLElement => {
    const found = host.querySelector<HTMLElement>('.fc-chart-settings');
    if (!found) throw new Error('missing chart settings');
    return found;
  };

  const openTab = (tab: 'elements' | 'styles' | 'filters' = 'elements'): HTMLElement => {
    host.querySelector<HTMLButtonElement>(`[data-chart-settings-tab="${tab}"]`)?.click();
    return surface();
  };

  const switchTab = (tab: 'elements' | 'styles' | 'filters'): HTMLElement => {
    surface().querySelector<HTMLButtonElement>(`#fc-chart-settings-tab-${tab}`)?.click();
    return surface();
  };

  const apply = (): void => {
    surface().querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.click();
  };

  const cancel = (): void => {
    surface().querySelector<HTMLButtonElement>('.fc-chart-settings__cancel')?.click();
  };

  const titleInput = (): HTMLInputElement => {
    const input = surface().querySelector<HTMLInputElement>(
      '.fc-chart-settings__field input[type="text"]',
    );
    if (!input) throw new Error('missing chart title input');
    return input;
  };

  const chartType = (): HTMLSelectElement =>
    fieldControl<HTMLSelectElement>(surface(), 'Chart type', 'select');

  const scatterStyle = (): HTMLSelectElement =>
    fieldControl<HTMLSelectElement>(surface(), 'Scatter style', 'select');

  const current = (): SessionChart | undefined => store.getState().charts.charts[0];

  const lastPatch = (): SessionChartPatch =>
    projection.patches[projection.patches.length - 1] as SessionChartPatch;

  const panel = (): HTMLElement => {
    const found = host.querySelector<HTMLElement>('.fc-chart');
    if (!found) throw new Error('missing chart overlay');
    return found;
  };

  const pixels = (element: HTMLElement) => ({
    x: Number.parseFloat(element.style.left),
    y: Number.parseFloat(element.style.top),
    w: Number.parseFloat(element.style.width),
    h: Number.parseFloat(element.style.height),
  });

  const fixture = (name: string): Uint8Array =>
    new Uint8Array(readFileSync(join(root, 'tests/fixtures/charts', name)));

  it('turns the legend off in the workbook and undoes it', () => {
    put([
      ['Month', 'Sales'],
      ['Jan', 10],
    ]);
    create(range(0, 0, 1, 1), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();

    expect(checkbox(openTab(), 'Legend').checked).toBe(true);
    apply();
    expect(projection.patches).toHaveLength(0);
    expect(history.canUndo()).toBe(false);

    setChecked(checkbox(openTab(), 'Legend'), false);
    apply();
    expect(current()?.legend?.visible).toBe(false);
    expect(host.querySelectorAll('[data-chart-role="legend"]')).toHaveLength(0);
    expect(history.undo()).toBe(true);
    expect(current()?.legend?.visible).toBe(true);
    expect(host.querySelectorAll('[data-chart-role="legend"]')).not.toHaveLength(0);
    handle.detach();
  });

  it('keeps an implicit legend sparse for a single series and allows an explicit one', () => {
    project({ kind: 'column' });
    const handle = attach();

    expect(checkbox(openTab(), 'Legend').checked).toBe(false);
    apply();
    expect(projection.patches).toHaveLength(0);

    setChecked(checkbox(openTab(), 'Legend'), true);
    apply();
    expect(lastPatch().legend).toEqual({ visible: true });
    expect(host.querySelectorAll('[data-chart-role="legend"]')).not.toHaveLength(0);
    handle.detach();
  });

  it('opens a non-modal settings surface and applies labels with one undoable update', () => {
    put([
      ['Month', 'Sales'],
      ['Jan', 10],
    ]);
    create(range(0, 0, 1, 1), { kind: 'column', title: 'Old title' });
    const handle = attach();
    const settingsButton = host.querySelector<HTMLButtonElement>(
      '.fc-chart__button[data-chart-settings-tab="elements"]',
    );
    expect(settingsButton).toBeTruthy();
    settingsButton?.click();
    const settings = surface();
    expect(settings.getAttribute('role')).toBe('dialog');
    expect(settings.getAttribute('aria-modal')).toBe('false');
    expect(settings.parentElement).toBe(host);
    expect(settings.querySelectorAll('[role="tab"]')).toHaveLength(3);
    expect(titleInput().value).toBe('Old title');
    setInput(titleInput(), 'New title');
    apply();

    expect(current()?.title).toBe('New title');
    expect(history.undo()).toBe(true);
    expect(current()?.title).toBe('Old title');
    expect(history.redo()).toBe(true);
    expect(current()?.title).toBe('New title');
    handle.detach();
  });

  it('keeps filter drafts outside the chart root and commits selected series/categories', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
      ['Feb', 30, 40],
    ]);
    create(range(0, 0, 2, 2), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    const settings = openTab('filters');
    expect(settings.closest('.fc-charts')).toBeNull();
    const filterGroups = settings.querySelectorAll<HTMLElement>('.fc-chart-settings__group');
    expect(filterGroups).toHaveLength(2);
    const seriesChecks =
      filterGroups[0]?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
    const categoryChecks =
      filterGroups[1]?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]');
    expect(seriesChecks).toHaveLength(2);
    expect(categoryChecks).toHaveLength(2);
    seriesChecks?.[0]?.click();
    categoryChecks?.[1]?.click();
    apply();
    expect(current()?.filter).toEqual({
      hiddenSeriesIds: ['series-0'],
      hiddenCategoryIds: ['category-1'],
    });
    expect(host.querySelectorAll('[data-chart-role="bar"]')).toHaveLength(1);
    handle.detach();
  });

  it('keeps pie variation stable for title-only and no-op Apply', () => {
    put([
      ['A', 'B'],
      [1, 2],
    ]);
    create(range(0, 0, 1, 1), { kind: 'pie' });
    const handle = attach();
    const fills = (): string[] =>
      Array.from(host.querySelectorAll<SVGElement>('[data-chart-role="slice"]')).map(
        (slice) => slice.getAttribute('fill') ?? '',
      );
    const before = fills();
    openTab();
    setInput(titleInput(), 'Updated pie');
    apply();
    expect(fills()).toEqual(before);
    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    openTab();
    apply();
    expect(history.canUndo()).toBe(false);
    expect(fills()).toEqual(before);
    handle.detach();
  });

  it('applies series styles and cancels drafts with focus restoration', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'line' });
    const handle = attach();
    expect(host.querySelector('[data-chart-role="line"]')?.getAttribute('data-series-id')).toBe(
      'series-0',
    );
    const settings = openTab('styles');
    expect(
      settings.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')?.textContent,
    ).toBe('Styles');
    apply();
    expect(history.canUndo()).toBe(false);

    // A line series has no fill: its first colour control is the line colour.
    const target = openTab('styles').querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const labels = Array.from(target?.querySelectorAll('.fc-chart-settings__label') ?? []).map(
      (label) => label.textContent ?? '',
    );
    expect(labels.some((label) => label.startsWith('Fill color'))).toBe(false);
    const lineColor = target?.querySelector<HTMLInputElement>('input[type="color"]');
    if (!lineColor) throw new Error('missing line colour');
    setInput(lineColor, '#ff0000');
    apply();
    expect(current()?.seriesStyles?.[0]?.line?.color).toBe('#ff0000');
    expect(current()?.seriesStyles?.[0]?.fill).toBeUndefined();
    expect(host.querySelector('[data-chart-role="line"]')?.getAttribute('stroke')).toBe('#ff0000');

    const cancelStylesButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-settings-tab="styles"]',
    );
    openTab('styles');
    setInput(titleInput(), 'discarded');
    cancel();
    expect(current()?.title).toEqual({ auto: true });
    expect(document.activeElement).toBe(cancelStylesButton);

    const elementsButton = host.querySelector<HTMLButtonElement>(
      '[data-chart-settings-tab="elements"]',
    );
    openTab('elements');
    setInput(titleInput(), 'escaped draft');
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(current()?.title).toEqual({ auto: true });
    expect(document.activeElement).toBe(elementsButton);
    handle.detach();
  });

  it('shows the series colours the workbook stores in the palette and style controls', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
      ['Feb', 30, 40],
    ]);
    create(range(0, 0, 2, 2), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    const barFill = (series: number): string | null | undefined =>
      host
        .querySelector<SVGElement>(`[data-chart-role="bar"][data-series-index="${series}"]`)
        ?.getAttribute('fill');
    const fills = current()?.seriesStyles?.map((style) => style.fill?.color);
    expect(fills).toHaveLength(2);
    expect(barFill(0)).toBe(fills?.[0]);
    expect(barFill(1)).toBe(fills?.[1]);

    let settings = openTab('styles');
    const palette = settings.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[0];
    const paletteColors = Array.from(
      palette?.querySelectorAll<HTMLInputElement>('input[type="color"]') ?? [],
    ).map((input) => input.value);
    expect(paletteColors.slice(0, 2)).toEqual(fills);
    let target = settings.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    expect(target?.querySelector<HTMLInputElement>('input[type="color"]')?.value).toBe(barFill(0));
    const seriesSelect = target?.querySelectorAll<HTMLSelectElement>('select')[0];
    if (!seriesSelect) throw new Error('missing series selector');
    setSelect(seriesSelect, 'series-1');
    settings = surface();
    target = settings.querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    expect(target?.querySelector<HTMLInputElement>('input[type="color"]')?.value).toBe(barFill(1));

    // A palette replaces the series' own colours.
    settings.querySelector<HTMLButtonElement>('[data-palette="mono"]')?.click();
    apply();
    const mono = current()?.seriesStyles?.map((style) => style.fill?.color);
    expect(mono).not.toEqual(fills);
    expect(barFill(0)).toBe(mono?.[0]);
    handle.detach();
  });

  it('shows the inherited colour of a pie slice in its point controls', () => {
    put([
      ['A', 'B'],
      [1, 2],
    ]);
    create(range(0, 0, 1, 1), { kind: 'pie', seriesDirection: 'rows' });
    const handle = attach();
    const secondSlice = host.querySelector<SVGElement>(
      '[data-chart-role="slice"][data-point-index="1"]',
    );
    const pieTarget = openTab('styles').querySelectorAll<HTMLElement>(
      '.fc-chart-settings__group',
    )[1];
    const pointSelect = pieTarget?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (!pointSelect) throw new Error('missing point selector');
    setSelect(pointSelect, 'category-1');
    const pointTarget = surface().querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    expect(pointTarget?.querySelector<HTMLInputElement>('input[type="color"]')?.value).toBe(
      secondSlice?.getAttribute('fill'),
    );
    handle.detach();
  });

  it('writes a point style to the workbook as that point', () => {
    put([
      ['Month', 'Sales'],
      ['Jan', 10],
      ['Feb', 15],
      ['Mar', 8],
    ]);
    const chart = create(range(0, 0, 3, 1), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    let target = openTab('styles').querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const pointSelect = target?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (!pointSelect) throw new Error('missing point selector');
    setSelect(pointSelect, 'category-1');
    target = surface().querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const pointFill = target?.querySelector<HTMLInputElement>('input[type="color"]');
    if (!pointFill) throw new Error('missing point fill');
    setInput(pointFill, '#ff0000');
    apply();

    const engineSeries = wb.getChart(0, chart.nativeObjectId)?.plots[0]?.series[0];
    expect(engineSeries?.points.map((point) => point.idx)).toEqual([1]);
    expect(current()?.seriesStyles?.[0]?.points).toEqual([
      { pointId: 'category-1', fill: { color: '#ff0000' } },
    ]);
    expect(
      host
        .querySelector('[data-chart-role="bar"][data-point-id="category-1"]')
        ?.getAttribute('fill'),
    ).toBe('#ff0000');
    expect(history.undo()).toBe(true);
    expect(current()?.seriesStyles?.[0]?.points).toBeUndefined();
    handle.detach();
  });

  it('keeps point overrides sparse when editing inherited line and marker fields', () => {
    project({
      kind: 'line',
      seriesStyles: [
        {
          seriesId: 'series-0',
          fill: { color: '#ff0000', transparency: 25 },
          line: { color: '#00aa00', transparency: 30, width: 2 },
          marker: { visible: true, shape: 'circle', size: 8 },
        },
      ],
    });
    const handle = attach();
    let target = openTab('styles').querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const pointSelect = target?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (!pointSelect) throw new Error('missing point selector');
    setSelect(pointSelect, 'category-0');
    target = surface().querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    // A line point has no fill fields: its colour is the line's and marker's.
    const numericInputs = target?.querySelectorAll<HTMLInputElement>('input[type="number"]');
    const lineTransparency = numericInputs?.[0];
    const markerSize = numericInputs?.[2];
    if (!lineTransparency || !markerSize) {
      throw new Error('missing point style inputs');
    }
    setInput(lineTransparency, '50');
    setInput(markerSize, '12');
    apply();
    const style = current()?.seriesStyles?.[0];
    expect(style?.fill).toEqual({ color: '#ff0000', transparency: 25 });
    expect(style?.line).toEqual({ color: '#00aa00', transparency: 30, width: 2 });
    expect(style?.marker).toEqual({ visible: true, shape: 'circle', size: 8 });
    expect(style?.points?.[0]).toEqual({
      pointId: 'category-0',
      // A transparency carries the colour it applies to, as a workbook line does.
      line: { color: '#00aa00', transparency: 50 },
      marker: { size: 12 },
    });
    handle.detach();
  });

  it('offers every writable marker symbol and applies a new one', () => {
    project({
      kind: 'line',
      seriesStyles: [{ seriesId: 'series-0', marker: { visible: true, shape: 'circle' } }],
    });
    const handle = attach();
    const target = openTab('styles').querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
    const shapes = Array.from(target?.querySelectorAll<HTMLSelectElement>('select') ?? []).find(
      (select) => Array.from(select.options).some((option) => option.value === 'triangle'),
    );
    if (!shapes) throw new Error('missing marker shape picker');
    // A picture marker is read from workbooks only, so it is not offered.
    expect(Array.from(shapes.options).map((option) => option.value)).toEqual([
      'circle',
      'square',
      'diamond',
      'triangle',
      'x',
      'star',
      'plus',
      'dot',
      'dash',
    ]);
    expect(Array.from(shapes.options).map((option) => option.textContent)).toEqual([
      'Circle',
      'Square',
      'Diamond',
      'Triangle',
      'X',
      'Star',
      'Plus',
      'Dot',
      'Dash',
    ]);
    setSelect(shapes, 'star');
    apply();
    expect(current()?.seriesStyles?.[0]?.marker?.shape).toBe('star');
    handle.detach();
  });

  it('takes the office preset and uncoloured series from the workbook theme accents', () => {
    project({ kind: 'column' });
    const handle = attach();
    const office = (): string[] => {
      const button = openTab('styles').querySelector('[data-palette="office"]');
      return Array.from(
        button?.querySelectorAll<HTMLElement>('.fc-chart-settings__palette-swatch') ?? [],
      ).map((swatch) => swatch.style.backgroundColor);
    };
    // The default workbook carries the default Office theme.
    expect(office()).toEqual(DEFAULT_THEME_ACCENTS);
    const palette = surface()
      .querySelectorAll<HTMLElement>('.fc-chart-settings__group')[0]
      ?.querySelectorAll<HTMLInputElement>('input[type="color"]');
    expect(Array.from(palette ?? []).map((input) => input.value)).toEqual(
      automaticPalette().slice(0, 6),
    );
    cancel();

    const theme = wb.getTheme()?.colors ?? [];
    const accents = [0xff4472c4, 0xffed7d31, 0xffa5a5a5, 0xffffc000, 0xff5b9bd5, 0xff70ad47];
    expect(wb.setThemeColors([...theme.slice(0, 4), ...accents, ...theme.slice(10)])).toBe(true);
    expect(office()).toEqual(['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47']);
    handle.detach();
  });

  it('shows and resets an automatic title', () => {
    project({ kind: 'column', title: 'Sales' });
    const handle = attach();
    const header = (): string | null | undefined =>
      host.querySelector('.fc-chart [data-chart-role="title"]')?.textContent;
    const automatic = (): HTMLInputElement | undefined =>
      Array.from(surface().querySelectorAll<HTMLInputElement>('input[type="checkbox"]')).find(
        (input) => input.closest('.fc-chart-settings__check')?.textContent === 'Automatic title',
      );
    expect(header()).toBe('Sales');
    openTab();
    const toggle = automatic();
    if (!toggle) throw new Error('missing automatic title option');
    expect(toggle.checked).toBe(false);
    toggle.click();
    apply();
    expect(current()?.title).toEqual({ auto: true });
    // One series: Excel's automatic title is its name.
    expect(header()).toBe('Value');

    // Typing a title replaces the automatic one.
    openTab();
    expect(automatic()?.checked).toBe(true);
    setInput(titleInput(), 'Revenue');
    expect(automatic()?.checked).toBe(false);
    apply();
    expect(current()?.title).toBe('Revenue');
    handle.detach();
  });

  it('toggles the automatic axis title and lets typed text replace it', () => {
    project({ kind: 'column', yAxis: { title: 'Units' } });
    const handle = attach();
    const changed = openTab();
    const yAxis = chartAxisFieldset(changed, 'Y-axis');
    const checkbox = (): HTMLInputElement | undefined =>
      Array.from(yAxis.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')).find(
        (input) => input.closest('.fc-chart-settings__check')?.textContent === 'Automatic title',
      );
    const title = yAxis.querySelector<HTMLInputElement>('input[type="text"]');
    if (!checkbox() || !title) throw new Error('missing axis title controls');
    expect(checkbox()?.checked).toBe(false);
    checkbox()?.click();
    apply();
    expect(current()?.yAxis?.title).toEqual({ auto: true });

    const again = chartAxisFieldset(openTab(), 'Y-axis');
    const toggle = Array.from(
      again.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'),
    ).find(
      (input) => input.closest('.fc-chart-settings__check')?.textContent === 'Automatic title',
    );
    expect(toggle?.checked).toBe(true);
    const input = again.querySelector<HTMLInputElement>('input[type="text"]');
    if (!input) throw new Error('missing axis title input');
    setInput(input, 'Revenue');
    expect(toggle?.checked).toBe(false);
    apply();
    expect(current()?.yAxis?.title).toBe('Revenue');
    handle.detach();
  });

  it('titles an automatic chart of several series with the localised chart title', () => {
    project(
      { kind: 'column', title: { auto: true } },
      {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'series-0', name: 'North', values: [1] },
          { id: 'series-1', name: 'South', values: [2] },
        ],
      },
    );
    const handle = attach({ labels: { chartTitle: 'グラフ タイトル' } });
    expect(host.querySelector('.fc-chart [data-chart-role="title"]')?.textContent).toBe(
      'グラフ タイトル',
    );
    handle.detach();
  });

  it('toggles a smoothed line on a line series, not on a point', () => {
    project({ kind: 'line', seriesStyles: [{ seriesId: 'series-0' }] });
    const handle = attach();
    const smoothToggle = (): HTMLInputElement | undefined => {
      const target = surface().querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1];
      return Array.from(
        target?.querySelectorAll<HTMLInputElement>('input[type="checkbox"]') ?? [],
      ).find(
        (input) => input.closest('.fc-chart-settings__check')?.textContent === 'Smoothed line',
      );
    };
    openTab('styles');
    const toggle = smoothToggle();
    if (!toggle) throw new Error('missing smoothed line toggle');
    expect(toggle.checked).toBe(false);
    toggle.click();
    apply();
    expect(current()?.seriesStyles?.[0]?.smooth).toBe(true);
    expect(host.querySelector('[data-chart-role="line"]')?.tagName.toLowerCase()).toBe('path');

    openTab('styles');
    const pointSelect = surface()
      .querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1]
      ?.querySelectorAll<HTMLSelectElement>('select')[1];
    if (!pointSelect) throw new Error('missing point selector');
    setSelect(pointSelect, 'category-0');
    expect(smoothToggle()).toBeUndefined();
    handle.detach();
  });

  describe('series add and remove', () => {
    const seriesSelect = (): HTMLSelectElement =>
      fieldControl<HTMLSelectElement>(surface(), 'Series', 'select');
    const button = (className: string): HTMLButtonElement => {
      const found = surface().querySelector<HTMLButtonElement>(`.${className}`);
      if (!found) throw new Error(`missing button: ${className}`);
      return found;
    };
    const refInput = (field: string): HTMLInputElement | null =>
      surface().querySelector<HTMLInputElement>(`[data-series-ref="${field}"]`);
    const typeRef = (field: string, value: string): void => {
      const input = refInput(field);
      if (!input) throw new Error(`missing reference input: ${field}`);
      setInput(input, value);
    };
    const seriesNames = (): (string | undefined)[] =>
      current()?.data.series.map((series) => series.name) ?? [];

    beforeEach(() => {
      put([
        [null, 'North', 'South', 'West'],
        ['Jan', 10, 20, 30],
        ['Feb', 15, 25, 35],
      ]);
    });

    it('validates the references, adds the series, and selects it', () => {
      create(range(0, 0, 2, 2), { seriesDirection: 'columns' });
      const handle = attach();
      openTab('styles');
      expect(refInput('values')).toBeNull();
      button('fc-chart-settings__add-series').click();
      expect(refInput('categories')).not.toBeNull();
      expect(refInput('x')).toBeNull();

      // Values are required and every typed field must be a reference.
      typeRef('name', 'not a ref');
      button('fc-chart-settings__new-series-confirm').click();
      expect(refInput('values')?.getAttribute('aria-invalid')).toBe('true');
      expect(refInput('name')?.getAttribute('aria-invalid')).toBe('true');
      expect(surface().querySelector('.fc-chart-settings__error')?.textContent).toBe(
        'Enter a valid cell reference.',
      );
      expect(seriesNames()).toEqual(['North', 'South']);

      // A sheet the workbook does not have parses but cannot be resolved.
      typeRef('name', 'D1');
      typeRef('values', 'Nowhere!D2:D3');
      button('fc-chart-settings__new-series-confirm').click();
      expect(refInput('values')?.getAttribute('aria-invalid')).toBe('true');
      expect(history.canUndo()).toBe(false);

      typeRef('values', 'Sheet1!$D$2:$D$3');
      typeRef('categories', 'A2:A3');
      button('fc-chart-settings__new-series-confirm').click();
      expect(seriesNames()).toEqual(['North', 'South', 'West']);
      expect(current()?.data).toMatchObject({ categories: ['Jan', 'Feb'] });
      expect(refInput('values')).toBeNull();
      expect(seriesSelect().value).toBe('series-2');
      expect(host.querySelectorAll('.fc-chart__plot [data-chart-role="bar"]')).toHaveLength(6);

      expect(history.undo()).toBe(true);
      expect(seriesNames()).toEqual(['North', 'South']);
      expect(seriesSelect().value).toBe('series-0');
      expect(history.redo()).toBe(true);
      expect(seriesNames()).toEqual(['North', 'South', 'West']);
      handle.detach();
    });

    it('asks a scatter series for X and Y values', () => {
      create(range(0, 1, 2, 2), { kind: 'scatter' });
      const handle = attach();
      openTab('styles');
      button('fc-chart-settings__add-series').click();
      expect(refInput('categories')).toBeNull();
      expect(fieldControl(surface(), 'Series Y values')).toBe(refInput('values'));
      typeRef('name', 'D1');
      typeRef('values', 'D2:D3');
      typeRef('x', 'B2:B3');
      button('fc-chart-settings__new-series-confirm').click();
      expect(seriesNames()).toContain('West');
      handle.detach();
    });

    it('removes the selected series and keeps the last one', () => {
      create(range(0, 0, 2, 2), { seriesDirection: 'columns' });
      const handle = attach();
      openTab('styles');
      setSelect(seriesSelect(), 'series-1');
      button('fc-chart-settings__remove-series').click();
      expect(seriesNames()).toEqual(['North']);
      expect(seriesSelect().value).toBe('series-0');

      const remove = button('fc-chart-settings__remove-series');
      expect(remove.disabled).toBe(true);
      expect(remove.title).toBe('A chart needs at least one series.');
      expect(button('fc-chart-settings__add-series').disabled).toBe(false);

      expect(history.undo()).toBe(true);
      expect(seriesNames()).toEqual(['North', 'South']);
      expect(button('fc-chart-settings__remove-series').disabled).toBe(false);
      handle.detach();
    });

    it('keeps a pending draft and drops its styles of a removed series', () => {
      create(range(0, 0, 2, 2), { seriesDirection: 'columns' });
      const handle = attach();
      openTab('styles');
      setSelect(seriesSelect(), 'series-1');
      const fill = surface()
        .querySelectorAll<HTMLElement>('.fc-chart-settings__group')[1]
        ?.querySelector<HTMLInputElement>('input[type="color"]');
      if (!fill) throw new Error('missing fill colour');
      setInput(fill, '#ff0000');
      button('fc-chart-settings__remove-series').click();
      switchTab('elements');
      setInput(titleInput(), 'Kept');
      apply();
      expect(current()?.title).toBe('Kept');
      expect(seriesNames()).toEqual(['North']);
      handle.detach();
    });
  });

  it('keeps settings drafts through a locale update and removes listeners on detach', () => {
    project({ kind: 'column' });
    const handle = attach({ locale: 'en-US' });
    const settings = openTab();
    setInput(titleInput(), 'draft title');
    handle.setLabels(
      {
        settings: 'グラフ設定',
        elementsTab: '要素',
        stylesTab: 'スタイル',
        filtersTab: 'フィルター',
      },
      'ja-JP',
    );
    expect(settings.lang).toBe('ja-JP');
    expect(settings.querySelector('.fc-chart-settings__title')?.textContent).toBe('グラフ設定');
    expect(titleInput().value).toBe('draft title');
    handle.detach();
    expect(host.querySelector('.fc-chart-settings')).toBeNull();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
  });

  it('reformats SVG data labels when the session locale changes', () => {
    put([[1234.5]]);
    create(range(0, 0, 0, 0), {
      kind: 'column',
      dataLabels: { visible: true, showValue: true },
    });
    const handle = attach({ locale: 'en-US', formatValue: formatSessionChartValue });
    const label = (): string | null =>
      host.querySelector<SVGTextElement>('[data-chart-role="data-label"]')?.textContent ?? null;
    expect(label()).toBe('1,234.5');
    handle.setLabels({}, 'de-DE');
    expect(label()).toBe('1.234,5');
    handle.detach();
  });

  it('uses Excel chart label positions and preserves an invalid current position', () => {
    const positionSelect = (): HTMLSelectElement => {
      const select = Array.from(
        surface().querySelectorAll<HTMLSelectElement>('.fc-chart-settings select'),
      ).find(
        (candidate) =>
          candidate.parentElement?.querySelector('.fc-chart-settings__label')?.textContent ===
            'Label position' &&
          !candidate.parentElement.classList.contains('fc-chart-settings__label-position'),
      );
      if (!select) throw new Error('missing chart label position select');
      return select;
    };
    const openPositionSelect = (
      kind: 'column' | 'line' | 'area',
      grouping?: 'clustered' | 'stacked',
      position?: 'outside-end',
    ): { handle: ReturnType<typeof attachSessionCharts>; select: HTMLSelectElement } => {
      project({
        kind,
        ...(grouping ? { grouping } : {}),
        dataLabels: { visible: true, ...(position ? { position } : {}) },
      });
      const handle = attach();
      openTab();
      return { handle, select: positionSelect() };
    };

    put([[10]]);
    create(range(0, 0, 0, 0), { kind: 'pie', dataLabels: { visible: true } });
    const pieHandle = attach();
    openTab();
    const pie = positionSelect();
    expect(pie.querySelector<HTMLOptionElement>('option[value="inside-end"]')?.disabled).toBe(
      false,
    );
    setSelect(pie, 'inside-end');
    apply();
    expect(current()?.dataLabels?.position).toBe('inside-end');
    pieHandle.detach();
    store = createSpreadsheetStore();

    const stacked = openPositionSelect('column', 'stacked', 'outside-end');
    expect(
      stacked.select.querySelector<HTMLOptionElement>('option[value="outside-end"]')?.disabled,
    ).toBe(true);
    stacked.handle.detach();

    const line = openPositionSelect('line');
    expect(line.select.querySelector('option[value="best-fit"]')?.textContent).toBe('Automatic');
    expect(['best-fit', 'right']).toContain(line.select.value);
    line.handle.detach();

    const area = openPositionSelect('area');
    expect(area.select.disabled).toBe(true);
    expect(area.select.querySelector<HTMLOptionElement>('option[value="best-fit"]')?.disabled).toBe(
      true,
    );
    area.handle.detach();
  });

  it('uses the built-in Excel formatter for direct chart attachment', () => {
    put([[1234.5]]);
    create(range(0, 0, 0, 0), {
      kind: 'column',
      dataLabels: { visible: true, showValue: true, numberFormat: '#,##0.00' },
    });
    const handle = attachSessionCharts({ host, store, locale: 'en-US' });
    expect(host.querySelector<SVGTextElement>('[data-chart-role="data-label"]')?.textContent).toBe(
      '1,234.50',
    );
    handle.detach();
  });

  it('patches only changed data-label fields and preserves sparse API values', () => {
    project({
      kind: 'column',
      dataLabels: { visible: true, showValue: true, color: '#123456', fontSize: 72 },
    });
    const handle = attach();
    const dataLabelFields = Array.from(openTab().querySelectorAll<HTMLElement>('fieldset')).find(
      (fieldset) => fieldset.querySelector('legend')?.textContent === 'Data labels',
    );
    if (!dataLabelFields) throw new Error('missing data label fields');
    setInput(fieldControl(dataLabelFields, 'Number format'), '#,##0.00');
    apply();
    expect(lastPatch()).toEqual({
      dataLabels: {
        visible: true,
        showValue: true,
        color: '#123456',
        fontSize: 72,
        numberFormat: '#,##0.00',
      },
    });
    handle.detach();
  });

  it('keeps implicit value labels when another label flag is enabled', () => {
    put([
      ['Month', 'Sales'],
      ['Jan', 10],
    ]);
    create(range(0, 0, 1, 1), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    const settings = openTab();
    const findCheckbox = (label: string): HTMLInputElement | undefined =>
      Array.from(settings.querySelectorAll<HTMLElement>('.fc-chart-settings__check'))
        .find((row) => row.textContent?.includes(label))
        ?.querySelector<HTMLInputElement>('input') ?? undefined;
    const showDataLabels = findCheckbox('Show data labels');
    const showCategoryName = findCheckbox('Category name');
    if (!showDataLabels || !showCategoryName) throw new Error('missing data-label controls');
    setChecked(showDataLabels, true);
    setChecked(showCategoryName, true);
    const dataLabelFields = Array.from(settings.querySelectorAll<HTMLElement>('fieldset')).find(
      (fieldset) => fieldset.querySelector('legend')?.textContent === 'Data labels',
    );
    if (!dataLabelFields) throw new Error('missing data label fields');
    setInput(fieldControl(dataLabelFields, 'Number format'), '#,##0.00');
    apply();
    expect(current()?.dataLabels?.showValue).toBe(true);
    expect(host.querySelector<SVGTextElement>('[data-chart-role="data-label"]')?.textContent).toBe(
      'Jan, 10.00',
    );
    handle.detach();
  });

  it('rejects settings Apply on a protected chart source', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'column', title: 'Original' });
    const handle = attach();
    openTab();
    setInput(titleInput(), 'Blocked');
    setProtectedSheet(store, 0, true);
    apply();
    expect(current()?.title).toBe('Original');
    handle.detach();
  });

  it('updates open settings editability when protection changes', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'column', title: 'Original' });
    const handle = attach();
    openTab();
    setInput(titleInput(), 'Draft');
    setProtectedSheet(store, 0, true);
    expect(titleInput().disabled).toBe(true);
    expect(titleInput().value).toBe('Draft');
    setProtectedSheet(store, 0, false);
    expect(titleInput().disabled).toBe(false);
    expect(titleInput().value).toBe('Draft');
    cancel();
    expect(current()?.title).toBe('Original');
    handle.detach();
  });

  it('keeps the settings surface read-only without a workbook', () => {
    project({ kind: 'column', title: 'Original' });
    const handle = attachSessionCharts({ host, store });
    openTab();
    expect(titleInput().disabled).toBe(true);
    expect(surface().querySelector<HTMLButtonElement>('.fc-chart-settings__apply')?.disabled).toBe(
      true,
    );
    handle.detach();
  });

  it('updates open settings editability when interaction policy changes', () => {
    const controller = new InteractionController({ store, history, getWb: () => wb });
    const unregister = registerInteractionController(store, controller);
    try {
      put([[1, 2]]);
      create(range(0, 0, 0, 1), { kind: 'column', title: 'Original' });
      const handle = attach();
      openTab();
      setInput(titleInput(), 'Draft');
      const applyButton = (): HTMLButtonElement | null =>
        surface().querySelector<HTMLButtonElement>('.fc-chart-settings__apply');
      expect(applyButton()?.disabled).toBe(false);

      controller.setPolicy(viewerPolicy());
      expect(applyButton()?.disabled).toBe(true);
      expect(titleInput().value).toBe('Draft');

      controller.setPolicy(undefined);
      expect(applyButton()?.disabled).toBe(false);
      expect(titleInput().value).toBe('Draft');
      handle.detach();
      expect(() => controller.setPolicy(viewerPolicy())).not.toThrow();
      expect(host.querySelector('.fc-chart-settings')).toBeNull();
    } finally {
      unregister();
      controller.dispose();
    }
  });

  it('shows a read-only settings surface for viewer policy', () => {
    const controller = new InteractionController({ store, history, getWb: () => wb });
    const unregister = registerInteractionController(store, controller);
    try {
      put([[1, 2]]);
      create(range(0, 0, 0, 1), { kind: 'column', title: 'Original' });
      const handle = attach();
      controller.setPolicy(viewerPolicy());
      openTab();
      expect(titleInput().disabled).toBe(true);
      const applyButton = surface().querySelector<HTMLButtonElement>('.fc-chart-settings__apply');
      expect(applyButton?.disabled).toBe(true);
      applyButton?.click();
      expect(current()?.title).toBe('Original');
      handle.detach();
    } finally {
      unregister();
      controller.dispose();
    }
  });

  it('applies gap width and overlap updates to the displayed columns', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
    ]);
    const chart = create(range(0, 0, 1, 2), {
      kind: 'column',
      seriesDirection: 'columns',
      gapWidth: 150,
      overlap: 0,
    });
    const handle = attach();
    const before = Array.from(host.querySelectorAll('[data-chart-role="bar"]'));
    expect(before).toHaveLength(2);
    const width = Number(before[0]?.getAttribute('width'));
    expect(width).toBeGreaterThan(0);
    expect(Number(before[1]?.getAttribute('x'))).toBeCloseTo(
      Number(before[0]?.getAttribute('x')) + width,
    );
    updateSessionChart(store, chart.id, { gapWidth: 0, overlap: 100 }, null, wb);
    const after = Array.from(host.querySelectorAll('[data-chart-role="bar"]'));
    expect(after).toHaveLength(2);
    expect(Number(after[0]?.getAttribute('width'))).toBeCloseTo(width * 3.5);
    expect(after[0]?.getAttribute('x')).toBe(after[1]?.getAttribute('x'));
    handle.detach();
  });

  it('edits axis fields sparsely, preserves unknown values, and accepts negative and zero values', () => {
    project({
      kind: 'column',
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
    const handle = attach();
    const axisInput = (fieldset: HTMLElement, text: string): HTMLInputElement => {
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

    openTab();
    apply();
    expect(projection.patches).toHaveLength(0);

    const changed = openTab();
    const xAxis = chartAxisFieldset(changed, 'X-axis');
    const yAxis = chartAxisFieldset(changed, 'Y-axis');
    setInput(axisInput(xAxis, 'X-axis title'), 'Updated X');
    setInput(axisInput(yAxis, 'Minimum'), '-5');
    setInput(axisInput(yAxis, 'Maximum'), '0');
    setInput(axisInput(yAxis, 'Major unit'), '0');
    apply();

    expect(projection.patches).toHaveLength(1);
    expect(current()?.xAxis).toMatchObject({ title: 'Updated X', customAxisValue: 'keep-x' });
    expect(current()?.yAxis).toMatchObject({
      min: -5,
      max: 0,
      numberFormat: '0.0',
      majorGridlines: false,
      customAxisValue: 'keep-y',
    });
    expect(current()?.yAxis?.majorUnit).toBeUndefined();
    handle.detach();
  });

  it('applies both axis reverse-order controls with undo, redo, and cancel', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'column' });
    const handle = attach();
    const before = { xAxis: current()?.xAxis, yAxis: current()?.yAxis };

    openTab();
    apply();
    expect(history.canUndo()).toBe(false);

    const noOp = openTab();
    setChecked(reverseOrder(noOp, 'X-axis'), true);
    setChecked(reverseOrder(noOp, 'X-axis'), false);
    setChecked(reverseOrder(noOp, 'Y-axis'), true);
    setChecked(reverseOrder(noOp, 'Y-axis'), false);
    apply();
    expect(current()?.xAxis).toEqual(before.xAxis);
    expect(history.canUndo()).toBe(false);

    const changed = openTab();
    setChecked(reverseOrder(changed, 'X-axis'), true);
    setChecked(reverseOrder(changed, 'Y-axis'), true);
    apply();
    expect(current()?.xAxis?.reverseOrder).toBe(true);
    expect(current()?.yAxis?.reverseOrder).toBe(true);

    expect(history.undo()).toBe(true);
    expect(current()?.xAxis).toEqual(before.xAxis);
    expect(current()?.yAxis).toEqual(before.yAxis);
    expect(history.redo()).toBe(true);
    expect(current()?.xAxis?.reverseOrder).toBe(true);
    expect(current()?.yAxis?.reverseOrder).toBe(true);

    setChecked(reverseOrder(openTab(), 'X-axis'), false);
    cancel();
    expect(current()?.xAxis?.reverseOrder).toBe(true);
    expect(current()?.yAxis?.reverseOrder).toBe(true);
    handle.detach();
  });

  it('syncs untouched external axis reversal while preserving a dirty draft', () => {
    project({ kind: 'column' });
    const handle = attach();
    const settings = openTab();
    setInput(titleInput(), 'Draft title');
    mutators.updateChart(store, projection.id, { xAxis: { reverseOrder: true } });
    expect(reverseOrder(settings, 'X-axis').checked).toBe(true);
    expect(titleInput().value).toBe('Draft title');

    setChecked(reverseOrder(settings, 'Y-axis'), true);
    mutators.updateChart(store, projection.id, { yAxis: { reverseOrder: false } });
    expect(reverseOrder(settings, 'Y-axis').checked).toBe(true);

    apply();
    expect(lastPatch().xAxis).toBeUndefined();
    expect(current()?.title).toBe('Draft title');
    expect(current()?.xAxis?.reverseOrder).toBe(true);
    expect(current()?.yAxis?.reverseOrder).toBe(true);
    handle.detach();
  });

  it.each(['pie', 'doughnut'] as const)('disables reverse-order controls for %s charts', (kind) => {
    project({ kind, xAxis: { reverseOrder: true }, yAxis: { reverseOrder: false } });
    const handle = attach();
    const settings = openTab();
    for (const axisLabel of ['X-axis', 'Y-axis']) {
      expect(reverseOrder(settings, axisLabel).disabled).toBe(true);
    }
    handle.detach();
  });

  it('disables axis controls according to chart roles and disables unsupported grouping', () => {
    const cases = [
      { kind: 'column' as const, xNumeric: false, yNumeric: true, grouping: false },
      { kind: 'bar' as const, xNumeric: true, yNumeric: false, grouping: false },
      { kind: 'scatter' as const, xNumeric: true, yNumeric: true, grouping: true },
      { kind: 'pie' as const, xNumeric: false, yNumeric: false, grouping: true },
    ];
    for (const item of cases) {
      store = createSpreadsheetStore();
      project({ kind: item.kind });
      const handle = attach();
      const settings = openTab();
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
      expect(minimum(chartAxisFieldset(settings, 'X-axis')).disabled).toBe(!item.xNumeric);
      expect(minimum(chartAxisFieldset(settings, 'Y-axis')).disabled).toBe(!item.yNumeric);
      expect(fieldControl<HTMLSelectElement>(settings, 'Grouping', 'select').disabled).toBe(
        item.grouping,
      );
      handle.detach();
    }
  });

  it('shows automatic for invalid stored major units and keeps the raw value until edited', () => {
    project({ kind: 'column', yAxis: { majorUnit: -1, customAxisValue: 'keep' } as never });
    const handle = attach();
    const yAxis = chartAxisFieldset(openTab(), 'Y-axis');
    expect(fieldControl(yAxis, 'Major unit').value).toBe('');
    setInput(fieldControl(yAxis, 'Y-axis title'), 'Value');
    apply();
    expect(current()?.yAxis).toMatchObject({
      majorUnit: -1,
      customAxisValue: 'keep',
      title: 'Value',
    });
    handle.detach();
  });

  it('shows forced zero for stacked line empty cells without rewriting the stored choice', () => {
    project({ kind: 'line', grouping: 'stacked', emptyCells: 'gap' });
    const handle = attach();
    const select = fieldControl<HTMLSelectElement>(openTab(), 'Empty cells', 'select');
    expect(select.disabled).toBe(true);
    expect(select.value).toBe('zero');
    apply();
    expect(projection.patches).toHaveLength(0);
    expect(current()?.emptyCells).toBe('gap');
    handle.detach();
  });

  it('enables scatter span only for line modes and preserves it through kind changes', () => {
    put([
      ['X', 'Sales'],
      [1, 10],
      [2, 20],
      [3, 30],
    ]);
    create(range(0, 0, 3, 1), { kind: 'scatter', seriesDirection: 'columns', emptyCells: 'gap' });
    const handle = attach();
    openTab();
    const emptyCells = (): HTMLSelectElement =>
      fieldControl<HTMLSelectElement>(surface(), 'Empty cells', 'select');
    const spanOption = (): HTMLOptionElement => {
      const option = Array.from(emptyCells().options).find(
        (candidate) => candidate.value === 'span',
      );
      if (!option) throw new Error('missing span option');
      return option;
    };

    expect(scatterStyle().value).toBe('marker');
    expect(spanOption().disabled).toBe(true);
    setSelect(scatterStyle(), 'line');
    expect(spanOption().disabled).toBe(false);
    setSelect(scatterStyle(), 'lineMarker');
    expect(spanOption().disabled).toBe(false);

    setSelect(scatterStyle(), 'line');
    setSelect(emptyCells(), 'span');
    setSelect(chartType(), 'column');
    expect(emptyCells().value).toBe('span');
    setSelect(chartType(), 'scatter');
    expect(scatterStyle().value).toBe('line');
    expect(emptyCells().value).toBe('span');
    expect(spanOption().disabled).toBe(false);

    apply();
    expect(current()).toMatchObject({ kind: 'scatter', emptyCells: 'span', scatterStyle: 'line' });
    // One segment between each pair of points.
    expect(host.querySelectorAll('[data-chart-role="line"]')).toHaveLength(2);
    handle.detach();
  });

  it('applies chart geometry knobs with Excel ranges and keeps pie angle separate', () => {
    put([[1, 2]]);
    const column = create(range(0, 0, 0, 1), { kind: 'column' });
    const handle = attach();
    const styles = openTab('styles');
    setInput(fieldControl(styles, 'Gap width'), '0');
    setInput(fieldControl(styles, 'Overlap'), '-100');
    apply();
    expect(current()).toMatchObject({ gapWidth: 0, overlap: -100 });
    handle.detach();
    mutators.removeChart(store, column.id);

    create(range(0, 0, 0, 1), { kind: 'pie' });
    const pieHandle = attach();
    setInput(fieldControl(openTab('styles'), 'First slice angle'), '360');
    apply();
    expect(current()?.firstSliceAngle).toBe(360);
    pieHandle.detach();
  });

  it('does not materialize clamped geometry defaults when another field changes', () => {
    project({ kind: 'column', gapWidth: 999, overlap: 200 });
    const handle = attach();
    openTab();
    setInput(titleInput(), 'Keep raw geometry');
    apply();
    expect(lastPatch()).toEqual({ title: 'Keep raw geometry' });
    handle.detach();
  });

  it('keeps an out-of-range doughnut hole raw until the field is edited', () => {
    project({ kind: 'doughnut', holeSize: 99 });
    const handle = attach();
    openTab();
    setInput(titleInput(), 'Keep raw hole');
    apply();
    expect(lastPatch()).toEqual({ title: 'Keep raw hole' });

    const hole = fieldControl(openTab('styles'), 'Hole size');
    expect(hole.value).toBe('90');
    setInput(hole, '90');
    apply();
    expect(current()?.holeSize).toBe(99);

    setInput(fieldControl(openTab('styles'), 'Hole size'), '70');
    apply();
    expect(lastPatch()).toEqual({ holeSize: 70 });
    handle.detach();
  });

  it('keeps a default doughnut hole sparse when edited back', () => {
    project({ kind: 'doughnut' });
    const handle = attach();
    setInput(fieldControl(openTab('styles'), 'Hole size'), '70');
    switchTab('elements');
    setInput(fieldControl(switchTab('styles'), 'Hole size'), '50');
    apply();
    expect(projection.patches).toHaveLength(0);
    handle.detach();
  });

  it('previews draft data across chart kinds and applies one undoable kind change', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 6],
      ['Feb', 15, 8],
      ['Mar', 8, 3],
    ]);
    create(range(0, 0, 3, 2), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    const chartBefore = structuredClone(current());
    const styleSelects = (): { series: HTMLSelectElement; points: HTMLSelectElement } => {
      const fields = Array.from(
        surface().querySelectorAll<HTMLElement>('.fc-chart-settings__field'),
      );
      const select = (label: string): HTMLSelectElement => {
        const found = fields
          .find((field) => field.querySelector('.fc-chart-settings__label')?.textContent === label)
          ?.querySelector<HTMLSelectElement>('select');
        if (!found) throw new Error(`missing ${label} select`);
        return found;
      };
      return { series: select('Series'), points: select('Point') };
    };
    const optionValues = (select: HTMLSelectElement): string[] =>
      Array.from(select.options).map((option) => option.value);

    openTab();
    setSelect(chartType(), 'scatter');
    cancel();
    expect(current()).toEqual(chartBefore);

    openTab();
    setSelect(chartType(), 'scatter');
    switchTab('styles');
    // Text categories become X values 1, 2, 3; every series stays.
    expect(Array.from(styleSelects().series.options).map((option) => option.textContent)).toEqual([
      'Sales',
      'Costs',
    ]);
    expect(optionValues(styleSelects().points)).toEqual([
      '__all__',
      'point-0',
      'point-1',
      'point-2',
    ]);

    switchTab('elements');
    setSelect(chartType(), 'doughnut');
    switchTab('styles');
    expect(styleSelects().series.options).toHaveLength(2);
    expect(optionValues(styleSelects().points)).toEqual([
      '__all__',
      'category-0',
      'category-1',
      'category-2',
    ]);
    apply();
    expect(current()?.kind).toBe('doughnut');
    expect(host.querySelectorAll('[data-chart-role="slice"]')).toHaveLength(6);
    expect(history.undo()).toBe(true);
    expect(current()?.kind).toBe('column');
    expect(history.redo()).toBe(true);
    expect(current()?.kind).toBe('doughnut');
    handle.detach();
  });

  it('previews the engine split when the draft changes the series direction', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 6],
      ['Feb', 15, 8],
      ['Mar', 8, 3],
    ]);
    create(range(0, 0, 3, 2), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    openTab();
    setSelect(fieldControl<HTMLSelectElement>(surface(), 'Series direction', 'select'), 'rows');
    const series = Array.from(
      switchTab('filters')
        .querySelectorAll<HTMLElement>('.fc-chart-settings__group')[0]
        ?.querySelectorAll<HTMLElement>('.fc-chart-settings__check') ?? [],
    ).map((row) => row.textContent?.trim());
    expect(series).toEqual(['Jan', 'Feb', 'Mar']);
    apply();
    expect(current()?.seriesDirection).toBe('rows');
    expect(current()?.data.series.map((entry) => entry.name)).toEqual(['Jan', 'Feb', 'Mar']);
    handle.detach();
  });

  it('keeps legend and vary-colors sparse across draft toggles', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 6],
      ['Feb', 15, 8],
    ]);
    create(range(0, 0, 2, 2), { kind: 'pie', seriesDirection: 'columns' });
    const handle = attach();

    const legend = checkbox(openTab(), 'Legend');
    expect(legend.checked).toBe(true);
    setChecked(legend, false);
    setChecked(legend, true);
    apply();
    expect(history.canUndo()).toBe(false);

    const vary = checkbox(openTab('styles'), 'Vary colors by point');
    expect(vary.checked).toBe(true);
    setChecked(vary, false);
    setChecked(vary, true);
    apply();
    expect(history.canUndo()).toBe(false);

    openTab();
    setSelect(chartType(), 'doughnut');
    const explicitFalse = checkbox(switchTab('styles'), 'Vary colors by point');
    expect(explicitFalse.checked).toBe(true);
    setChecked(explicitFalse, false);
    apply();
    expect(current()).toMatchObject({ kind: 'doughnut', varyColors: false });
    expect(history.canUndo()).toBe(true);
    handle.detach();
  });

  it('keeps optional defaults sparse when a new kind checkbox is toggled back', () => {
    project({ kind: 'column' });
    const handle = attach();
    expect(checkbox(openTab(), 'Legend').checked).toBe(false);
    setSelect(chartType(), 'doughnut');

    const legend = checkbox(surface(), 'Legend');
    expect(legend.checked).toBe(true);
    setChecked(legend, false);
    setChecked(legend, true);

    const vary = checkbox(switchTab('styles'), 'Vary colors by point');
    expect(vary.checked).toBe(true);
    setChecked(vary, false);
    setChecked(vary, true);
    switchTab('elements');
    setSelect(chartType(), 'column');
    apply();
    expect(projection.patches).toHaveLength(0);
    handle.detach();
  });

  it('keeps an explicit vary-colors value when toggled back', () => {
    project({ kind: 'column', varyColors: false });
    const handle = attach();
    const vary = checkbox(openTab('styles'), 'Vary colors by point');
    expect(vary.checked).toBe(false);
    setChecked(vary, true);
    setChecked(vary, false);
    apply();
    expect(projection.patches).toHaveLength(0);
    handle.detach();
  });

  it('does not overwrite an external kind update when applying another draft field', () => {
    project({ kind: 'column' });
    const handle = attach();
    openTab();
    setInput(titleInput(), 'External kind survives');
    mutators.updateChart(store, projection.id, { kind: 'scatter' });
    apply();
    expect(lastPatch()).toEqual({ title: 'External kind survives' });
    expect(current()).toMatchObject({ kind: 'scatter', title: 'External kind survives' });
    handle.detach();
  });

  it('keeps external legend and vary-colors overrides when applying only a title', () => {
    project({ kind: 'column', legend: { visible: false }, varyColors: false });
    const handle = attach();
    const openAndEditTitle = (value: string): HTMLElement => {
      const settings = openTab();
      setInput(titleInput(), value);
      return settings;
    };

    let settings = openAndEditTitle('External true');
    mutators.updateChart(store, projection.id, { legend: { visible: true }, varyColors: true });
    expect(checkbox(settings, 'Legend').checked).toBe(true);
    apply();
    expect(lastPatch()).toEqual({ title: 'External true' });
    expect(current()).toMatchObject({ legend: { visible: true }, varyColors: true });

    settings = openAndEditTitle('External false');
    settings = switchTab('styles');
    expect(checkbox(settings, 'Vary colors by point').checked).toBe(true);
    mutators.updateChart(store, projection.id, { legend: { visible: false }, varyColors: false });
    expect(checkbox(surface(), 'Vary colors by point').checked).toBe(false);
    apply();
    expect(lastPatch()).toEqual({ title: 'External false' });
    expect(current()).toMatchObject({ legend: { visible: false }, varyColors: false });
    handle.detach();
  });

  it('refreshes chart kind and axis roles for external same-domain kind changes', () => {
    project({ kind: 'column' });
    const handle = attach();
    openTab();
    const axisMinimum = (axis: 'X-axis' | 'Y-axis'): HTMLInputElement => {
      const field = Array.from(
        chartAxisFieldset(surface(), axis).querySelectorAll<HTMLElement>(
          '.fc-chart-settings__field',
        ),
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

    mutators.updateChart(store, projection.id, { kind: 'bar' });
    expect(chartType().value).toBe('bar');
    expect(axisMinimum('X-axis').disabled).toBe(false);
    expect(axisMinimum('Y-axis').disabled).toBe(true);

    mutators.updateChart(store, projection.id, { kind: 'line' });
    expect(chartType().value).toBe('line');
    expect(axisMinimum('X-axis').disabled).toBe(true);
    expect(axisMinimum('Y-axis').disabled).toBe(false);
    handle.detach();
  });

  it('shows all scatter styles and resolves marker visibility for each mode', () => {
    put([
      ['X', 'Sales'],
      [1, 10],
      [2, 20],
      [3, 30],
    ]);
    create(range(0, 0, 3, 1), { kind: 'scatter', seriesDirection: 'columns' });
    const handle = attach();
    openTab();
    expect(Array.from(scatterStyle().options).map((option) => option.value)).toEqual([
      'marker',
      'line',
      'lineMarker',
    ]);
    expect(scatterStyle().value).toBe('marker');

    const markerVisible = (): HTMLInputElement => {
      const row = Array.from(
        surface().querySelectorAll<HTMLElement>('.fc-chart-settings__check'),
      ).find((candidate) => candidate.textContent?.includes('Visible'));
      const input = row?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error('missing marker checkbox');
      return input;
    };
    setSelect(scatterStyle(), 'line');
    switchTab('styles');
    expect(markerVisible().checked).toBe(false);
    setChecked(markerVisible(), true);

    switchTab('elements');
    setSelect(scatterStyle(), 'lineMarker');
    switchTab('styles');
    expect(markerVisible().checked).toBe(true);
    apply();
    expect(current()?.seriesStyles?.[0]?.marker?.visible).toBe(true);
    handle.detach();
  });

  it('keeps scatter style sparse through kind cycles and preserves dormant edits', () => {
    put([
      ['X', 'Sales'],
      [1, 10],
    ]);
    create(range(0, 0, 1, 1), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    openTab();
    setSelect(chartType(), 'scatter');
    expect(scatterStyle().value).toBe('marker');
    setSelect(scatterStyle(), 'line');
    setSelect(chartType(), 'column');
    setSelect(chartType(), 'scatter');
    expect(scatterStyle().value).toBe('line');
    setSelect(scatterStyle(), 'marker');
    setSelect(chartType(), 'column');
    apply();
    expect(current()).toMatchObject({ kind: 'column' });
    expect(current()?.scatterStyle).toBeUndefined();
    expect(history.canUndo()).toBe(false);
    handle.detach();
  });

  it('preserves external scatter styles and no-op drafts', () => {
    project({ kind: 'scatter' });
    const handle = attach();
    openTab();
    setSelect(scatterStyle(), 'line');
    setSelect(scatterStyle(), 'marker');
    apply();
    expect(projection.patches).toHaveLength(0);

    openTab();
    mutators.updateChart(store, projection.id, { scatterStyle: 'line' });
    expect(scatterStyle().value).toBe('line');
    setInput(titleInput(), 'External line');
    mutators.updateChart(store, projection.id, { scatterStyle: 'lineMarker' });
    expect(scatterStyle().value).toBe('lineMarker');
    apply();
    expect(lastPatch()).toEqual({ title: 'External line' });
    expect(current()).toMatchObject({ title: 'External line', scatterStyle: 'lineMarker' });

    openTab();
    setSelect(scatterStyle(), 'marker');
    cancel();
    expect(current()?.scatterStyle).toBe('lineMarker');
    handle.detach();
  });

  it.each(['column', 'bar', 'line', 'area', 'pie', 'scatter'] as const)(
    'renders %s through the chart scene and shows the title once',
    (kind) => {
      put(
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
            ],
      );
      create(range(0, 0, 3, 1), { kind, title: 'Revenue', seriesDirection: 'columns' });
      const handle = attach();
      const marks = host.querySelectorAll('.fc-chart__plot [data-series-index="0"]');
      expect(marks.length).toBeGreaterThan(0);
      const titles = host.querySelectorAll('.fc-chart [data-chart-role="title"]');
      expect(titles).toHaveLength(1);
      expect(titles[0]?.textContent).toBe('Revenue');
      handle.detach();
    },
  );

  it('preserves distinct series and updates chart marks when source cells change', () => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 20],
      ['Feb', 30, 40],
    ]);
    create(range(0, 0, 2, 2), {
      kind: 'column',
      seriesDirection: 'columns',
      yAxis: { min: 0, max: 100 },
    });
    const handle = attach();
    expect(host.querySelectorAll('[data-chart-role="bar"]')).toHaveLength(4);
    expect(host.querySelectorAll('text[data-chart-role="legend"]')).toHaveLength(2);
    const selector = '[data-chart-role="bar"][data-series-index="0"][data-point-index="0"]';
    const before = Number(host.querySelector(selector)?.getAttribute('height'));
    const unsubscribe = wb.subscribe((event) => {
      if (event.kind === 'recalc') refreshChartsAfterRecalc(wb, store);
    });
    try {
      wb.setNumber({ sheet: 0, row: 1, col: 1 }, 50);
    } finally {
      unsubscribe();
    }
    const after = Number(host.querySelector(selector)?.getAttribute('height'));
    expect(before).toBeGreaterThan(0);
    expect(after).toBeCloseTo(before * 5);
    handle.detach();
  });

  it('renders chart overlays and deletes them from the workbook', () => {
    put([[10, 20]]);
    const chart = create(range(0, 0, 0, 1), { kind: 'column', title: 'Revenue' });
    const handle = attach();

    expect(host.querySelector('.fc-chart [data-chart-role="title"]')?.textContent).toBe('Revenue');
    expect(host.querySelector('svg rect[data-chart-role="bar"]')).toBeTruthy();
    // Like Excel, a chart has no close button: Delete removes the selected chart.
    expect(host.querySelector('.fc-chart__close')).toBeNull();
    host.tabIndex = -1;

    panel().dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }));
    expect(store.getState().charts.charts).toHaveLength(0);
    expect(wb.listDrawingObjects(0).map((object) => object.objectId)).not.toContain(
      chart.nativeObjectId,
    );
    expect(host.querySelector('.fc-chart')).toBeNull();
    expect(document.activeElement).toBe(host);
    expect(history.undo()).toBe(true);
    expect(store.getState().charts.charts.map((entry) => entry.id)).toEqual([chart.id]);
    handle.detach();
  });

  it('uses localized fallback labels for untitled charts', () => {
    project({ kind: 'line' });
    const handle = attach({
      labels: {
        columnChart: '縦棒グラフ',
        lineChart: '折れ線グラフ',
        elementsButton: 'グラフ要素',
      },
    });

    // An untitled chart draws no title; its kind names it for assistive technology.
    expect(host.querySelector('.fc-chart [data-chart-role="title"]')).toBeNull();
    expect(panel().getAttribute('aria-label')).toBe('折れ線グラフ');
    const elements = (): HTMLButtonElement | null =>
      host.querySelector('.fc-chart__button[data-chart-settings-tab="elements"]');
    expect(elements()?.getAttribute('aria-label')).toBe('グラフ要素');
    expect(elements()?.title).toBe('グラフ要素');

    handle.setLabels({ lineChart: '線グラフ', elementsButton: '要素' });
    expect(panel().getAttribute('aria-label')).toBe('線グラフ');
    expect(elements()?.getAttribute('aria-label')).toBe('要素');
    handle.detach();
  });

  it('keeps bar and column marks inside padded slots for sparse and dense series', () => {
    // 334 x 146 px leaves the 320 x 132 scene inside the frame's border and padding.
    const sceneAnchor = { ...PROJECTED_ANCHOR, cx: 334 * 9525, cy: 146 * 9525 };
    for (const kind of ['bar', 'column'] as const) {
      for (const count of [1, 2, 16, 100]) {
        store = createSpreadsheetStore();
        project(
          { kind, nativeAnchor: sceneAnchor },
          {
            domain: 'category',
            categories: Array.from({ length: count }, (_, index) => `C${index + 1}`),
            series: [
              {
                id: 'series-0',
                values: Array.from({ length: count }, (_, index) => index + 1),
              },
            ],
          },
        );
        const handle = attach();
        expect(host.querySelector('.fc-chart__plot')?.getAttribute('viewBox')).toBe('0 0 320 132');
        const marks = Array.from(
          host.querySelectorAll<SVGRectElement>('.fc-chart__plot rect[data-chart-role="bar"]'),
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
          const mark = marks[index];
          if (!previous || !mark) throw new Error('chart mark missing from measured series');
          expect(mark.start).toBeGreaterThanOrEqual(previous.start + previous.size - 1e-9);
        }
        handle.detach();
      }
    }
  });

  it('places the frame at the chart anchor and follows it', () => {
    put([[1, 2]]);
    const chart = create(range(0, 0, 0, 1), { kind: 'column' });
    const handle = attach();
    const before = pixels(panel());
    // Excel's default frame is 360 x 216 pt.
    expect(before.w).toBeCloseTo(480);
    expect(before.h).toBeCloseTo(288);
    const { fromRow, fromRowOff } = chart.nativeAnchor;
    updateSessionChart(
      store,
      chart.id,
      { anchor: { row: fromRow + 2, rowOffEmu: fromRowOff } },
      null,
      wb,
    );
    expect(pixels(panel()).y).toBeCloseTo(before.y + 40);
    handle.detach();
  });

  it('routes chart object edits through the protection guard', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'line' });
    const handle = attach();
    const plot = host.querySelector<SVGSVGElement>('.fc-chart__plot');
    if (!plot) throw new Error('missing chart plot');
    setProtectedSheet(store, 0, true);
    const before = structuredClone(current());
    const start = pixels(panel());

    firePointer(plot, 'pointerdown', { clientX: 10, clientY: 10 });
    firePointer(window, 'pointermove', { clientX: 45, clientY: 60 });
    firePointer(window, 'pointerup', { clientX: 45, clientY: 60 });
    fireKey(panel(), 'ArrowRight');
    fireKey(panel(), 'Delete');

    expect(store.getState().charts.charts).toHaveLength(1);
    expect(current()).toEqual(before);
    expect(pixels(panel())).toEqual(start);
    setProtectedSheet(store, 0, false);
    fireKey(panel(), 'ArrowRight');
    expect(pixels(panel()).x).toBeCloseTo(start.x + 8);
    handle.detach();
  });

  it('moves and resizes chart overlays through pointer interactions', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'line', title: 'Trend' });
    const handle = attach();
    const start = pixels(panel());
    // Like Excel, the chart moves from a press anywhere on it.
    const plot = host.querySelector<SVGSVGElement>('.fc-chart__plot');
    if (!plot) throw new Error('missing chart plot');
    firePointer(plot, 'pointerdown', { clientX: 10, clientY: 10 });
    firePointer(window, 'pointermove', { clientX: 45, clientY: 60 });
    firePointer(window, 'pointerup', { clientX: 45, clientY: 60 });
    const moved = pixels(panel());
    expect(moved.x).toBeCloseTo(start.x + 35);
    expect(moved.y).toBeCloseTo(start.y + 50);

    expect(host.querySelectorAll('.fc-drawing-handle')).toHaveLength(8);
    const resize = host.querySelector<HTMLElement>('.fc-drawing-handle[data-handle="se"]');
    if (!resize) throw new Error('missing chart resize handle');
    firePointer(resize, 'pointerdown', { clientX: 100, clientY: 100 });
    firePointer(window, 'pointermove', { clientX: 180, clientY: 145 });
    firePointer(window, 'pointerup', { clientX: 180, clientY: 145 });
    const resized = pixels(panel());
    expect(resized.w).toBeCloseTo(start.w + 80);
    expect(resized.h).toBeCloseTo(start.h + 45);
    expect(history.undo()).toBe(true);
    expect(pixels(panel()).w).toBeCloseTo(start.w);
    handle.detach();
  });

  it('marks the active chart overlay as selected and brings it forward', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'column', title: 'First' });
    create(range(0, 0, 0, 1), { kind: 'line', title: 'Second' });
    const handle = attach();
    const charts = Array.from(host.querySelectorAll<HTMLElement>('.fc-chart'));
    expect(charts).toHaveLength(2);
    const [first, second] = charts;
    if (!first || !second) throw new Error('missing chart overlays');
    // The workbook's stacking order: the later chart sits on top.
    expect(Number(second.style.zIndex)).toBeGreaterThan(Number(first.style.zIndex));

    firePointer(first, 'pointerdown');
    expect(first.classList.contains('fc-chart--selected')).toBe(true);
    expect(first.getAttribute('aria-selected')).toBe('true');
    expect(Number(first.style.zIndex)).toBeGreaterThan(Number(second.style.zIndex));
    expect(second.getAttribute('aria-selected')).toBe('false');

    second.focus();
    expect(second.classList.contains('fc-chart--selected')).toBe(true);
    expect(Number(second.style.zIndex)).toBeGreaterThan(Number(first.style.zIndex));
    expect(first.getAttribute('aria-selected')).toBe('false');
    handle.detach();
  });

  it('moves, resizes, and removes focused chart overlays through keyboard interactions', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'column', title: 'Accessible chart' });
    const handle = attach();
    expect(panel().tabIndex).toBe(0);
    expect(panel().getAttribute('aria-label')).toBe('Accessible chart');
    expect(panel().getAttribute('aria-roledescription')).toBe('chart');
    expect(panel().getAttribute('aria-keyshortcuts')).toContain('Shift+ArrowDown');
    const start = pixels(panel());

    const move = fireKey(panel(), 'ArrowRight');
    expect(move.defaultPrevented).toBe(true);
    expect(pixels(panel()).x).toBeCloseTo(start.x + 8);
    expect(pixels(panel()).y).toBeCloseTo(start.y);

    const resize = fireKey(panel(), 'ArrowDown', { shiftKey: true });
    expect(resize.defaultPrevented).toBe(true);
    expect(pixels(panel()).w).toBeCloseTo(start.w);
    expect(pixels(panel()).h).toBeCloseTo(start.h + 8);

    fireKey(panel(), 'Delete');
    expect(store.getState().charts.charts).toHaveLength(0);
    expect(host.querySelector('.fc-chart')).toBeNull();
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('draws a chart it cannot render as a placeholder that only moves, resizes and deletes', async () => {
    const loaded = await WorkbookHandle.loadBytes(fixture('chart_chartex.xlsx'));
    try {
      hydrateChartsFromEngine(loaded, store, 0);
      const handle = attachSessionCharts({
        host,
        store,
        history,
        getWb: () => loaded,
        labels: { unsupportedChart: 'このグラフの種類は表示できません' },
      });
      const placeholder = panel();
      expect(placeholder.classList.contains('fc-chart--unsupported')).toBe(true);
      expect(host.querySelector('.fc-chart [data-chart-role="title"]')?.textContent).toBe('cx1');
      expect(host.querySelector('[data-chart-settings-tab]')).toBeNull();
      expect(
        host.querySelector('.fc-chart__plot [data-chart-role="unsupported"]')?.textContent,
      ).toBe('このグラフの種類は表示できません');
      expect(host.querySelectorAll('.fc-drawing-handle')).toHaveLength(8);

      const start = pixels(placeholder);
      fireKey(placeholder, 'ArrowDown');
      expect(pixels(panel()).y).toBeCloseTo(start.y + 8);
      fireKey(panel(), 'Delete');
      expect(store.getState().charts.charts).toHaveLength(0);
      expect(history.undo()).toBe(true);
      expect(current()?.kind).toBe('unsupported');
      handle.detach();
    } finally {
      loaded.dispose();
    }
  });

  it('disables the series and data range settings of a PivotChart', async () => {
    const loaded = await WorkbookHandle.loadBytes(fixture('chart_pivot.xlsx'));
    try {
      hydrateChartsFromEngine(loaded, store, 1);
      mutators.setSheetIndex(store, 1);
      const handle = attachSessionCharts({ host, store, history, getWb: () => loaded });
      expect(current()?.pivotChart).toBe(true);
      expect(
        host.querySelectorAll('.fc-chart__plot [data-chart-role="bar"]').length,
      ).toBeGreaterThan(0);
      const direction = fieldControl<HTMLSelectElement>(openTab(), 'Series direction', 'select');
      expect(direction.disabled).toBe(true);
      expect(fieldControl<HTMLSelectElement>(surface(), 'Chart type', 'select').disabled).toBe(
        false,
      );
      const filters = Array.from(
        switchTab('filters').querySelectorAll<HTMLInputElement | HTMLButtonElement>(
          '.fc-chart-settings__body input, .fc-chart-settings__body button',
        ),
      );
      expect(filters.length).toBeGreaterThan(0);
      expect(filters.every((control) => control.disabled)).toBe(true);
      const styles = switchTab('styles');
      for (const action of ['add-series', 'remove-series']) {
        const control = styles.querySelector<HTMLButtonElement>(`.fc-chart-settings__${action}`);
        expect(control?.disabled).toBe(true);
        expect(control?.title).toBe(
          'The series and data range of a PivotChart follow its PivotTable.',
        );
      }
      handle.detach();
    } finally {
      loaded.dispose();
    }
  });

  const salesTable = (): SessionChart => {
    put([
      ['Month', 'Sales', 'Costs'],
      ['Jan', 10, 6],
      ['Feb', 15, 8],
    ]);
    return create(range(0, 0, 2, 2), { kind: 'column', seriesDirection: 'columns' });
  };

  const drag = (target: EventTarget, dx: number, dy: number, init: PointerEventInit = {}) => {
    firePointer(target, 'pointerdown', { clientX: 100, clientY: 100, ...init });
    firePointer(window, 'pointermove', { clientX: 100 + dx, clientY: 100 + dy, ...init });
    firePointer(window, 'pointerup', { clientX: 100 + dx, clientY: 100 + dy, ...init });
  };

  it('resizes from any sizing handle and keeps the shape of a shift-dragged corner', () => {
    salesTable();
    const handle = attach();
    const grip = (name: string): HTMLElement => {
      const found = host.querySelector<HTMLElement>(`.fc-drawing-handle[data-handle="${name}"]`);
      if (!found) throw new Error(`missing ${name} handle`);
      return found;
    };
    const start = pixels(panel());

    // A leading corner moves the origin with it.
    drag(grip('nw'), -20, -10);
    let frame = pixels(panel());
    expect(frame.x).toBeCloseTo(start.x - 20);
    expect(frame.y).toBeCloseTo(start.y - 10);
    expect(frame.w).toBeCloseTo(start.w + 20);
    expect(frame.h).toBeCloseTo(start.h + 10);

    // An edge handle changes one dimension only.
    const before = pixels(panel());
    drag(grip('e'), 40, 30);
    frame = pixels(panel());
    expect(frame.w).toBeCloseTo(before.w + 40);
    expect(frame.h).toBeCloseTo(before.h);
    expect(frame.x).toBeCloseTo(before.x);

    const ratio = frame.w / frame.h;
    drag(grip('se'), 60, 0, { shiftKey: true });
    const scaled = pixels(panel());
    expect(scaled.w).toBeCloseTo(frame.w + 60);
    expect(scaled.w / scaled.h).toBeCloseTo(ratio, 2);
    handle.detach();
  });

  it('selects a chart on a click without moving it', () => {
    salesTable();
    const handle = attach();
    const start = pixels(panel());
    const plot = host.querySelector<SVGSVGElement>('.fc-chart__plot');
    if (!plot) throw new Error('missing chart plot');
    firePointer(plot, 'pointerdown', { clientX: 50, clientY: 50 });
    firePointer(window, 'pointermove', { clientX: 51, clientY: 51 });
    firePointer(window, 'pointerup', { clientX: 51, clientY: 51 });
    expect(panel().classList.contains('fc-chart--selected')).toBe(true);
    expect(pixels(panel())).toEqual(start);
    expect(history.undo()).toBe(false);
    handle.detach();
  });

  it('leaves the chart on a press elsewhere or on Escape', () => {
    salesTable();
    const handle = attach();
    firePointer(panel(), 'pointerdown');
    expect(panel().classList.contains('fc-chart--selected')).toBe(true);
    firePointer(host, 'pointerdown');
    expect(panel().classList.contains('fc-chart--selected')).toBe(false);
    expect(panel().getAttribute('aria-selected')).toBe('false');

    host.tabIndex = -1;
    host.focus();
    panel().focus();
    expect(panel().classList.contains('fc-chart--selected')).toBe(true);
    const escaped = fireKey(panel(), 'Escape');
    expect(escaped.defaultPrevented).toBe(true);
    expect(panel().classList.contains('fc-chart--selected')).toBe(false);
    expect(document.activeElement).toBe(host);
    handle.detach();
  });

  it('opens the chart settings on a double click', () => {
    salesTable();
    const handle = attach();
    panel().dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
    const settings = surface();
    expect(settings.hidden).toBe(false);
    expect(settings.querySelector('[role="tab"][aria-selected="true"]')?.id).toBe(
      'fc-chart-settings-tab-elements',
    );
    handle.detach();
  });

  it('labels each data point the way Excel words its tooltip', () => {
    salesTable();
    const handle = attach({
      labels: { pointTooltip: '系列 "{series}" 要素 "{point}"', pointValue: '値: {value}' },
    });
    const tooltip = (series: number, point: number): string | null | undefined =>
      host.querySelector(
        `[data-chart-role="bar"][data-series-index="${series}"][data-point-index="${point}"] title`,
      )?.textContent;
    expect(tooltip(0, 0)).toBe('系列 "Sales" 要素 "Jan"\n値: 10');
    expect(tooltip(1, 1)).toBe('系列 "Costs" 要素 "Feb"\n値: 8');
    handle.detach();
  });

  it('gives a pie slice its share of the total', () => {
    project(
      { kind: 'pie' },
      {
        domain: 'category',
        categories: ['A', 'B'],
        categoryIds: ['category-0', 'category-1'],
        series: [{ id: 'series-0', name: 'Value', values: [30, 10] }],
      },
    );
    const handle = attach();
    const slice = host.querySelector('[data-chart-role="slice"][data-point-index="0"] title');
    expect(slice?.textContent).toBe('Series "Value" Point "A"\nValue: 30 (75%)');
    handle.detach();
  });

  it('outlines the selected chart data by part and clears it when the chart is left', () => {
    salesTable();
    const handle = attach();
    expect(store.getState().ui.chartSourceRefs).toEqual([]);
    firePointer(panel(), 'pointerdown');
    expect(store.getState().ui.chartSourceRefs).toEqual([
      { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2, role: 'name' },
      { sheet: 0, r0: 1, c0: 0, r1: 2, c1: 0, role: 'categories' },
      { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2, role: 'values' },
    ]);
    // Selecting keeps the overlay in place, so focus and pointer capture survive.
    const selected = panel();
    expect(selected.isConnected).toBe(true);
    firePointer(host, 'pointerdown');
    expect(store.getState().ui.chartSourceRefs).toEqual([]);
    firePointer(panel(), 'pointerdown');
    handle.detach();
    expect(store.getState().ui.chartSourceRefs).toEqual([]);
    expect(store.getState().ui.selectedObject).toBeNull();
  });

  const fireContextMenu = (target: EventTarget, clientX = 120, clientY = 140): MouseEvent => {
    const e = new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX, clientY });
    target.dispatchEvent(e);
    return e;
  };
  const objectMenu = (): HTMLElement | null =>
    document.querySelector<HTMLElement>('.fc-ctxmenu--object');
  const menuItems = (): { id: string | undefined; label: string | null; disabled: boolean }[] =>
    Array.from(objectMenu()?.querySelectorAll<HTMLButtonElement>('.fc-ctxmenu__item') ?? []).map(
      (item) => ({ id: item.dataset.fcAction, label: item.textContent, disabled: item.disabled }),
    );

  it('opens the chart menu on a right click instead of the cell menu', () => {
    salesTable();
    const second = create(range(0, 0, 2, 2), { kind: 'line', seriesDirection: 'columns' });
    const handle = attach({ labels: { objectName: 'グラフ {n}' } });
    const hostMenu = vi.fn();
    host.addEventListener('contextmenu', hostMenu);
    const [back, front] = Array.from(host.querySelectorAll<HTMLElement>('.fc-chart'));
    if (!back || !front) throw new Error('missing charts');

    const event = fireContextMenu(back);
    expect(event.defaultPrevented).toBe(true);
    expect(hostMenu).not.toHaveBeenCalled();
    expect(objectMenu()?.style.display).toBe('block');
    expect(objectMenu()?.getAttribute('aria-label')).toBe('グラフ 1');
    expect(menuItems()).toEqual([
      { id: 'changeChartType', label: 'Change Chart Type...', disabled: false },
      { id: 'selectData', label: 'Select Data...', disabled: false },
      { id: 'bringToFront', label: 'Bring to Front', disabled: false },
      { id: 'sendToBack', label: 'Send to Back', disabled: true },
      { id: 'formatChartArea', label: 'Format Chart Area...', disabled: false },
    ]);
    expect(store.getState().ui.selectedObject?.name).toBe('グラフ 1');
    expect(document.activeElement?.textContent).toBe('Change Chart Type...');

    objectMenu()?.querySelector<HTMLButtonElement>('[data-fc-action="bringToFront"]')?.click();
    expect(objectMenu()?.style.display).toBe('none');
    expect(
      store
        .getState()
        .charts.charts.map((chart) => chart.id)
        .at(-1),
    ).not.toBe(second.id);
    expect(wb.listDrawingObjects(0).at(-1)?.objectId).not.toBe(second.nativeObjectId);
    expect(history.undo()).toBe(true);
    expect(
      store
        .getState()
        .charts.charts.map((chart) => chart.id)
        .at(-1),
    ).toBe(second.id);

    host.removeEventListener('contextmenu', hostMenu);
    handle.detach();
    expect(objectMenu()).toBeNull();
  });

  it('opens settings from the chart menu and closes the menu with Escape', () => {
    salesTable();
    const handle = attach();
    fireKey(panel(), 'F10', { shiftKey: true });
    expect(objectMenu()?.style.display).toBe('block');
    const escaped = fireKey(document.activeElement ?? panel(), 'Escape');
    expect(escaped.defaultPrevented).toBe(true);
    expect(objectMenu()?.style.display).toBe('none');
    expect(document.activeElement).toBe(panel());
    expect(panel().classList.contains('fc-chart--selected')).toBe(true);

    fireContextMenu(panel());
    objectMenu()?.querySelector<HTMLButtonElement>('[data-fc-action="selectData"]')?.click();
    expect(host.querySelector('.fc-chart-settings')).not.toBeNull();
    expect(
      host.querySelector('#fc-chart-settings-tab-filters')?.getAttribute('aria-selected'),
    ).toBe('true');
    expect(panel().classList.contains('fc-chart--selected')).toBe(true);
    handle.detach();
  });

  it('names the selected chart for the name box and releases it when left', () => {
    salesTable();
    const handle = attach({ labels: { objectName: 'グラフ {n}' } });
    expect(store.getState().ui.selectedObject).toBeNull();
    firePointer(panel(), 'pointerdown');
    expect(store.getState().ui.selectedObject).toEqual({
      id: panel().dataset.chartId,
      name: 'グラフ 1',
    });
    firePointer(host, 'pointerdown');
    expect(store.getState().ui.selectedObject).toBeNull();
    handle.detach();
  });

  it("leaves another overlay's selected object alone", () => {
    salesTable();
    const handle = attach();
    mutators.setSelectedObject(store, { id: 'picture-1', name: 'Picture 1' });
    firePointer(host, 'pointerdown');
    expect(store.getState().ui.selectedObject).toEqual({ id: 'picture-1', name: 'Picture 1' });
    handle.detach();
    expect(store.getState().ui.selectedObject).toEqual({ id: 'picture-1', name: 'Picture 1' });
  });

  it('scrolls a chart off with its anchor and hides it under the headers', () => {
    salesTable();
    const handle = attach();
    const start = pixels(panel());
    expect(panel().style.clipPath).toBe('');
    // B2 anchors the chart; scrolling three columns past it carries the chart off the leading edge.
    const anchorCol = current()?.nativeAnchor.fromCol ?? 0;
    mutators.scrollBy(store, 0, anchorCol + 3);
    const scrolled = pixels(panel());
    expect(scrolled.x).toBeLessThan(start.x);
    expect(scrolled.x).toBeLessThan(0);
    expect(panel().style.clipPath).toMatch(/^inset\(-64px -64px -64px \d+(\.\d+)?px\)$/);
    mutators.scrollBy(store, 0, -(anchorCol + 3));
    expect(pixels(panel())).toEqual(start);
    expect(panel().style.clipPath).toBe('');
    handle.detach();
  });

  it('keeps chart button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/session-charts.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain("className: 'fc-chart__button'");
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps session charts close to Excel 365 desktop object geometry', () => {
    const css = readFileSync(
      join(root, 'src/styles/core/app/overlays/quick-analysis-and-charts.css'),
      'utf8',
    );
    const chartCss = css.slice(css.indexOf('.fc-chart {'));

    // A flat chart area: a hairline frame, no rounding and no shadow.
    const frame = chartCss.slice(0, chartCss.indexOf('}'));
    expect(frame).not.toContain('border-radius');
    expect(frame).not.toContain('box-shadow');
    // Round sizing handles on every edge and corner, shared with pictures and
    // shown only while selected.
    const drawingCss = readFileSync(
      join(root, 'src/styles/core/app/overlays/drawings.css'),
      'utf8',
    );
    expect(drawingCss).toMatch(
      /\.fc-drawing-handle\s*\{[\s\S]*?display: none;[\s\S]*?width: 9px;[\s\S]*?height: 9px;[\s\S]*?border-radius: 50%;/,
    );
    expect(drawingCss).toMatch(
      /\.fc-chart--selected \.fc-drawing-handle,\s*\.fc-illustration--selected \.fc-drawing-handle\s*\{\s*display: block;/,
    );
    // The elements, styles and filters buttons stand outside the chart's right edge.
    expect(chartCss).toMatch(
      /\.fc-chart__buttons\s*\{[\s\S]*?left: calc\(100% \+ 8px\);[\s\S]*?flex-direction: column;/,
    );
    expect(chartCss).toMatch(/\.fc-chart__button\s*\{[\s\S]*?width: 26px;[\s\S]*?height: 26px;/);
    expect(chartCss).not.toContain('.fc-chart__header');
    expect(chartCss).not.toContain('.fc-chart__close');
  });

  it('keeps sparse crossing auto and invalid raw values untouched until edited', () => {
    project({ kind: 'column' });
    const handle = attach();

    let settings = openTab();
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('auto');
    apply();
    expect(projection.patches).toHaveLength(0);

    settings = openTab();
    setSelect(chartAxisCrossingSelect(settings, 'X-axis'), 'min');
    setSelect(chartAxisCrossingSelect(surface(), 'X-axis'), 'auto');
    apply();
    expect(projection.patches).toHaveLength(0);

    mutators.updateChart(store, projection.id, {
      xAxis: { crossing: Number.NaN } as never,
      yAxis: { crossing: Number.POSITIVE_INFINITY } as never,
    });
    settings = openTab();
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('auto');
    expect(chartAxisCrossingSelect(settings, 'Y-axis').value).toBe('auto');
    setInput(titleInput(), 'Raw title');
    apply();
    expect(lastPatch()).toEqual({ title: 'Raw title' });
    expect(Number.isNaN(current()?.xAxis?.crossing as number)).toBe(true);
    expect(current()?.yAxis?.crossing).toBe(Number.POSITIVE_INFINITY);

    mutators.updateChart(store, projection.id, {
      xAxis: { crossing: 'auto' },
      yAxis: { crossing: 'auto' },
    });
    const applied = projection.patches.length;
    settings = openTab();
    setSelect(chartAxisCrossingSelect(settings, 'X-axis'), 'min');
    setSelect(chartAxisCrossingSelect(surface(), 'X-axis'), 'auto');
    apply();
    expect(projection.patches).toHaveLength(applied);
    expect(current()?.xAxis?.crossing).toBe('auto');
    handle.detach();
  });

  it('normalizes category crossings, keeps numeric defaults, and supports cancel', () => {
    project(
      { kind: 'column', filter: { hiddenCategoryIds: ['category-1'] } },
      {
        domain: 'category',
        categories: ['A', 'B', 'C'],
        categoryIds: ['category-0', 'category-1', 'category-2'],
        series: [{ id: 'series-0', name: 'Value', values: [1, 2, 3] }],
      },
    );
    const handle = attach();

    let settings = openTab();
    setSelect(chartAxisCrossingSelect(settings, 'X-axis'), 'custom');
    const xCustom = chartAxisCrossingCustom(surface(), 'X-axis');
    expect(xCustom.min).toBe('1');
    expect(xCustom.step).toBe('1');
    setInput(xCustom, '2.6');
    setSelect(chartAxisCrossingSelect(surface(), 'Y-axis'), 'custom');
    settings = surface();
    expect(chartAxisCrossingCustom(settings, 'X-axis').value).toBe('2');
    const yCustom = chartAxisCrossingCustom(settings, 'Y-axis');
    expect(yCustom.step).toBe('any');
    expect(yCustom.value).toBe('0');
    setInput(yCustom, '-2.5');
    apply();
    expect(current()?.xAxis?.crossing).toBe(2);
    expect(current()?.yAxis?.crossing).toBe(-2.5);

    settings = openTab();
    setSelect(chartAxisCrossingSelect(settings, 'X-axis'), 'min');
    setSelect(chartAxisCrossingSelect(surface(), 'X-axis'), 'custom');
    expect(chartAxisCrossingCustom(surface(), 'X-axis').value).toBe('2');
    cancel();

    setSelect(chartAxisCrossingSelect(openTab(), 'X-axis'), 'min');
    apply();
    expect(current()?.xAxis?.crossing).toBe('min');

    setSelect(chartAxisCrossingSelect(openTab(), 'X-axis'), 'max');
    apply();
    expect(current()?.xAxis?.crossing).toBe('max');

    setSelect(chartAxisCrossingSelect(openTab(), 'X-axis'), 'auto');
    apply();
    expect(current()?.xAxis?.crossing).toBeUndefined();

    setSelect(chartAxisCrossingSelect(openTab(), 'X-axis'), 'min');
    cancel();
    expect(current()?.xAxis?.crossing).toBeUndefined();
    handle.detach();
  });

  it('applies a custom axis crossing in the workbook with undo and redo', () => {
    put([
      ['Category', 'Value'],
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ]);
    create(range(0, 0, 3, 1), { kind: 'column', seriesDirection: 'columns' });
    const handle = attach();
    setSelect(chartAxisCrossingSelect(openTab(), 'Y-axis'), 'custom');
    setInput(chartAxisCrossingCustom(surface(), 'Y-axis'), '-2.5');
    apply();
    expect(current()?.yAxis?.crossing).toBe(-2.5);
    expect(history.undo()).toBe(true);
    expect(current()?.yAxis?.crossing).toBe('auto');
    expect(history.redo()).toBe(true);
    expect(current()?.yAxis?.crossing).toBe(-2.5);
    handle.detach();
  });

  it('retains dormant circular crossings and maps numeric roles per chart family', () => {
    project({ kind: 'pie', xAxis: { crossing: 2 }, yAxis: { crossing: 'max' } });
    const handle = attach();
    const settings = openTab();
    expect(chartAxisCrossingSelect(settings, 'X-axis').disabled).toBe(true);
    expect(chartAxisCrossingSelect(settings, 'Y-axis').disabled).toBe(true);
    expect(chartAxisCrossingCustom(settings, 'X-axis').disabled).toBe(true);
    apply();
    expect(projection.patches).toHaveLength(0);
    handle.detach();
  });

  it('syncs untouched crossings and preserves a dirty crossing with other draft fields', () => {
    project({ kind: 'column' });
    const handle = attach();
    let settings = openTab();
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('auto');
    mutators.updateChart(store, projection.id, { xAxis: { crossing: 'min' } });
    settings = surface();
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('min');

    setInput(titleInput(), 'Draft title');
    setSelect(chartAxisCrossingSelect(settings, 'X-axis'), 'custom');
    setInput(chartAxisCrossingCustom(surface(), 'X-axis'), '2');
    mutators.updateChart(store, projection.id, { xAxis: { crossing: 'max' } });
    settings = surface();
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('custom');
    expect(titleInput().value).toBe('Draft title');
    apply();
    expect(current()?.title).toBe('Draft title');
    expect(current()?.xAxis?.crossing).toBe(2);
    handle.detach();
  });

  it('keeps sparse label positions and invalid raw values until edited', () => {
    project({ kind: 'column' });
    const handle = attach();

    let settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('automatic');
    apply();
    expect(projection.patches).toHaveLength(0);

    mutators.updateChart(store, projection.id, { xAxis: { labelPosition: 'low' } });
    settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('low');
    apply();
    expect(projection.patches).toHaveLength(0);

    setSelect(chartAxisLabelPositionSelect(openTab(), 'X-axis'), 'automatic');
    apply();
    expect(projection.patches).toHaveLength(1);
    expect(current()?.xAxis?.labelPosition).toBeUndefined();

    settings = openTab();
    setSelect(chartAxisLabelPositionSelect(settings, 'X-axis'), 'nextTo');
    setSelect(chartAxisLabelPositionSelect(surface(), 'X-axis'), 'automatic');
    apply();
    expect(projection.patches).toHaveLength(1);

    mutators.updateChart(store, projection.id, {
      xAxis: { labelPosition: 'invalid-runtime-value' } as never,
      yAxis: { labelPosition: 17 } as never,
    });
    settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('automatic');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('automatic');
    setInput(titleInput(), 'Unrelated title');
    apply();
    expect(lastPatch()).toEqual({ title: 'Unrelated title' });
    expect(current()?.xAxis?.labelPosition).toBe('invalid-runtime-value');
    expect(current()?.yAxis?.labelPosition).toBe(17);

    setSelect(chartAxisLabelPositionSelect(openTab(), 'X-axis'), 'high');
    cancel();
    expect(current()?.xAxis?.labelPosition).toBe('invalid-runtime-value');
    handle.detach();
  });

  it('applies an axis label position in the workbook with undo and redo', () => {
    put([[1, 2]]);
    create(range(0, 0, 0, 1), { kind: 'column' });
    const handle = attach();
    setSelect(chartAxisLabelPositionSelect(openTab(), 'X-axis'), 'low');
    apply();
    expect(current()?.xAxis?.labelPosition).toBe('low');
    expect(history.undo()).toBe(true);
    expect(current()?.xAxis?.labelPosition).toBe('nextTo');
    expect(history.redo()).toBe(true);
    expect(current()?.xAxis?.labelPosition).toBe('low');
    handle.detach();
  });

  it('keeps dormant circular label positions through chart kind cycles', () => {
    project({ kind: 'pie', xAxis: { labelPosition: 'high' }, yAxis: { labelPosition: 'none' } });
    const handle = attach();

    let settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('high');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('none');
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').disabled).toBe(true);
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').disabled).toBe(true);
    apply();
    expect(projection.patches).toHaveLength(0);

    mutators.updateChart(store, projection.id, { kind: 'column' });
    settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').disabled).toBe(false);
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('high');
    mutators.updateChart(store, projection.id, { kind: 'pie' });
    settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('high');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('none');
    apply();
    expect(projection.patches).toHaveLength(0);
    handle.detach();
  });

  it('syncs untouched label positions while protecting a dirty position and title', () => {
    project({ kind: 'column' });
    const handle = attach();
    let settings = openTab();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('automatic');

    mutators.updateChart(store, projection.id, {
      xAxis: { labelPosition: 'low' },
      yAxis: { labelPosition: 'high' },
    });
    settings = surface();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('low');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('high');

    setSelect(chartAxisCrossingSelect(settings, 'X-axis'), 'min');
    mutators.updateChart(store, projection.id, {
      xAxis: { crossing: 'max', labelPosition: 'high' },
      yAxis: { labelPosition: 'low' },
    });
    settings = surface();
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('min');
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('high');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('low');

    setInput(titleInput(), 'Draft title');
    setSelect(chartAxisLabelPositionSelect(settings, 'X-axis'), 'nextTo');
    mutators.updateChart(store, projection.id, {
      xAxis: { crossing: 'max', labelPosition: 'none' },
      yAxis: { labelPosition: 'high' },
    });
    settings = surface();
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('nextTo');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('high');
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('min');
    expect(titleInput().value).toBe('Draft title');

    apply();
    expect(current()?.title).toBe('Draft title');
    expect(current()?.xAxis?.labelPosition).toBe('nextTo');
    expect(current()?.xAxis?.crossing).toBe('min');
    expect(current()?.yAxis?.labelPosition).toBe('high');
    handle.detach();
  });

  it('keeps sparse logarithmic settings, remembered bases, and invalid raw values', () => {
    project({ kind: 'column', yAxis: { majorUnit: 2 } });
    const handle = attach();
    const majorUnit = (settings: HTMLElement): HTMLInputElement => {
      const field = Array.from(
        chartAxisFieldset(settings, 'Y-axis').querySelectorAll<HTMLElement>(
          '.fc-chart-settings__field',
        ),
      ).find(
        (candidate) =>
          candidate.querySelector('.fc-chart-settings__label')?.textContent === 'Major unit',
      );
      const input = field?.querySelector<HTMLInputElement>('input');
      if (!input) throw new Error('missing major unit input');
      return input;
    };

    let settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'Y-axis').checked).toBe(false);
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('10');
    expect(chartAxisLogBase(settings, 'Y-axis').min).toBe('2');
    expect(chartAxisLogBase(settings, 'Y-axis').max).toBe('1000');
    expect(chartAxisLogBase(settings, 'Y-axis').step).toBe('any');
    expect(chartAxisLogBase(settings, 'Y-axis').disabled).toBe(true);
    expect(majorUnit(settings).disabled).toBe(false);

    setChecked(chartAxisLogCheckbox(settings, 'Y-axis'), true);
    settings = surface();
    expect(chartAxisLogBase(settings, 'Y-axis').disabled).toBe(false);
    expect(majorUnit(settings).disabled).toBe(true);
    setInput(chartAxisLogBase(settings, 'Y-axis'), '20');
    setChecked(chartAxisLogCheckbox(settings, 'Y-axis'), false);
    settings = surface();
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('20');
    expect(chartAxisLogBase(settings, 'Y-axis').disabled).toBe(true);
    expect(majorUnit(settings).disabled).toBe(false);
    setChecked(chartAxisLogCheckbox(settings, 'Y-axis'), true);
    settings = surface();
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('20');
    setChecked(chartAxisLogCheckbox(settings, 'Y-axis'), false);
    apply();
    expect(projection.patches).toHaveLength(0);
    expect(current()?.yAxis?.majorUnit).toBe(2);

    settings = openTab();
    setChecked(chartAxisLogCheckbox(settings, 'Y-axis'), true);
    setInput(chartAxisLogBase(surface(), 'Y-axis'), '1001');
    apply();
    expect(current()?.yAxis?.logBase).toBe(1000);

    setInput(chartAxisLogBase(openTab(), 'Y-axis'), '');
    apply();
    expect(current()?.yAxis?.logBase).toBe(10);

    mutators.updateChart(store, projection.id, { yAxis: { logBase: 1 } as never });
    settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'Y-axis').checked).toBe(false);
    setInput(titleInput(), 'Unrelated title');
    apply();
    expect(lastPatch()).toEqual({ title: 'Unrelated title' });
    expect(current()?.yAxis?.logBase).toBe(1);

    const applied = projection.patches.length;
    settings = openTab();
    setChecked(chartAxisLogCheckbox(settings, 'Y-axis'), true);
    setChecked(chartAxisLogCheckbox(surface(), 'Y-axis'), false);
    apply();
    expect(projection.patches).toHaveLength(applied);
    expect(current()?.yAxis?.logBase).toBe(1);

    setChecked(chartAxisLogCheckbox(openTab(), 'Y-axis'), true);
    cancel();
    expect(current()?.yAxis?.logBase).toBe(1);
    handle.detach();
  });

  it('applies a logarithmic value axis in the workbook with undo and redo', () => {
    put([[1, 100]]);
    create(range(0, 0, 0, 1), { kind: 'column' });
    const handle = attach();
    setChecked(chartAxisLogCheckbox(openTab(), 'Y-axis'), true);
    setInput(chartAxisLogBase(surface(), 'Y-axis'), '20');
    apply();
    expect(current()?.yAxis?.logBase).toBe(20);
    expect(history.undo()).toBe(true);
    expect(current()?.yAxis?.logBase).toBeUndefined();
    expect(history.redo()).toBe(true);
    expect(current()?.yAxis?.logBase).toBe(20);
    handle.detach();
  });

  it('keeps dormant logarithmic bases by chart type and grouping', () => {
    project({
      kind: 'column',
      grouping: 'stacked',
      xAxis: { logBase: 10 },
      yAxis: { logBase: 5 },
    });
    const handle = attach();

    let settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'X-axis').checked).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'X-axis').disabled).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').checked).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').disabled).toBe(true);
    expect(chartAxisLogBase(settings, 'Y-axis').disabled).toBe(true);
    apply();
    expect(projection.patches).toHaveLength(0);

    mutators.updateChart(store, projection.id, { grouping: 'percent-stacked' });
    settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'Y-axis').checked).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').disabled).toBe(true);

    mutators.updateChart(store, projection.id, { kind: 'bar', grouping: 'clustered' });
    settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'X-axis').checked).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'X-axis').disabled).toBe(false);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').checked).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').disabled).toBe(true);

    mutators.updateChart(store, projection.id, { kind: 'scatter' });
    settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'X-axis').disabled).toBe(false);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').disabled).toBe(false);
    expect(chartAxisLogBase(settings, 'X-axis').value).toBe('10');
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('5');

    mutators.updateChart(store, projection.id, { kind: 'pie' });
    settings = openTab();
    expect(chartAxisLogCheckbox(settings, 'X-axis').disabled).toBe(true);
    expect(chartAxisLogCheckbox(settings, 'Y-axis').disabled).toBe(true);
    apply();
    expect(projection.patches).toHaveLength(0);
    handle.detach();
  });

  it('syncs untouched logarithmic axes around dirty crossing and label fields', () => {
    project({ kind: 'column' });
    const handle = attach();

    let settings = openTab();
    mutators.updateChart(store, projection.id, { yAxis: { logBase: 20 } });
    settings = surface();
    expect(chartAxisLogCheckbox(settings, 'Y-axis').checked).toBe(true);
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('20');

    setSelect(chartAxisCrossingSelect(settings, 'Y-axis'), 'min');
    mutators.updateChart(store, projection.id, { yAxis: { crossing: 'max', logBase: 30 } });
    settings = surface();
    expect(chartAxisCrossingSelect(settings, 'Y-axis').value).toBe('min');
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('30');

    setSelect(chartAxisLabelPositionSelect(settings, 'Y-axis'), 'nextTo');
    mutators.updateChart(store, projection.id, {
      yAxis: { crossing: 'max', labelPosition: 'high', logBase: 40 },
    });
    settings = surface();
    expect(chartAxisCrossingSelect(settings, 'Y-axis').value).toBe('min');
    expect(chartAxisLabelPositionSelect(settings, 'Y-axis').value).toBe('nextTo');
    expect(chartAxisLogBase(settings, 'Y-axis').value).toBe('40');

    mutators.updateChart(store, projection.id, { kind: 'scatter' });
    settings = openTab();
    setChecked(chartAxisLogCheckbox(settings, 'X-axis'), true);
    mutators.updateChart(store, projection.id, {
      xAxis: { crossing: 'max', labelPosition: 'none', logBase: 50 },
    });
    settings = surface();
    expect(chartAxisLogCheckbox(settings, 'X-axis').checked).toBe(true);
    expect(chartAxisLogBase(settings, 'X-axis').value).toBe('10');
    expect(chartAxisCrossingSelect(settings, 'X-axis').value).toBe('max');
    expect(chartAxisLabelPositionSelect(settings, 'X-axis').value).toBe('none');

    apply();
    expect(current()?.xAxis?.logBase).toBe(10);
    expect(current()?.xAxis?.crossing).toBe('max');
    expect(current()?.xAxis?.labelPosition).toBe('none');
    handle.detach();
  });
});
