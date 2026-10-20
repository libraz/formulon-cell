import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { chartUsesPalette } from '../../../../src/commands/chart-design.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createInsertDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/insert.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators, type SessionChart } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('chart tabs of the ribbon', () => {
  let workbook: WorkbookHandle;
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ locale: 'en', workbook });
    const rows = [
      [null, 'North', 'South'],
      ['Jan', 10, 4],
      ['Feb', 15, 6],
      ['Mar', 12, 9],
    ];
    rows.forEach((cells, row) => {
      cells.forEach((value, col) => {
        if (typeof value === 'number') workbook.setNumber({ sheet: 0, row, col }, value);
        else if (typeof value === 'string') workbook.setText({ sheet: 0, row, col }, value);
      });
    });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
    workbook.dispose();
  });

  const tabIds = (): string[] =>
    Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-tab]')).map(
      (tab) => tab.dataset.ribbonTab ?? '',
    );
  const command = (id: string): HTMLButtonElement => {
    const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`);
    if (!found) throw new Error(`missing ${id}`);
    return found;
  };
  const onlyChart = (): SessionChart => {
    const [chart] = sheet.instance.store.getState().charts.charts;
    if (!chart) throw new Error('no chart');
    return chart;
  };
  /** Press a cell, which leaves a selected chart. */
  const pressSheet = (): void => {
    sheet.host.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
  };
  const insertChart = (): SessionChart => {
    createInsertDropdownDefaults(sheet.instance).createChartFromSelection('column');
    return onlyChart();
  };

  it('opens Chart Design on an inserted chart and goes back once the chart is left', () => {
    const onTabChange = vi.fn();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
      onTabChange,
    });
    tb.setActiveTab('insert');
    expect(tabIds()).not.toContain('chartDesign');

    insertChart();
    expect(tabIds().slice(-2)).toEqual(['chartDesign', 'chartFormat']);
    expect(tb.getActiveTab()).toBe('chartDesign');
    expect(onTabChange).toHaveBeenLastCalledWith('chartDesign');
    const design = host.querySelector<HTMLElement>('[data-ribbon-tab="chartDesign"]');
    expect(design?.classList.contains('fc-tb__ribbon-tab--contextual')).toBe(true);
    expect(design?.getAttribute('aria-selected')).toBe('true');

    // A press on the ribbon keeps the chart selected.
    command('chartSwitchRowColumn').dispatchEvent(
      new PointerEvent('pointerdown', { bubbles: true }),
    );
    expect(sheet.instance.store.getState().ui.selectedObject?.id).toBe(onlyChart().id);

    pressSheet();
    expect(tabIds()).not.toContain('chartDesign');
    expect(tb.getActiveTab()).toBe('insert');
    expect(onTabChange).toHaveBeenLastCalledWith('insert');
    tb.dispose();
  });

  it('shows the chart tabs, without opening them, for a chart selected on the sheet', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const chart = insertChart();
    pressSheet();
    tb.setActiveTab('home');

    mutators.setSelectedObject(sheet.instance.store, { id: chart.id, name: 'Chart 1' });
    expect(tabIds()).toContain('chartFormat');
    expect(tb.getActiveTab()).toBe('home');

    // Leaving a chart tab the user opened goes back to the tab before it.
    tb.setActiveTab('chartFormat');
    pressSheet();
    expect(tb.getActiveTab()).toBe('home');
    tb.dispose();
  });

  it('edits the selected chart from the Chart Design commands', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    insertChart();
    expect(onlyChart().data.series.map((series) => series.name)).toEqual(['North', 'South']);

    command('chartSwitchRowColumn').click();
    expect(onlyChart().data.series.map((series) => series.name)).toEqual(['Jan', 'Feb', 'Mar']);

    command('chartAddElement').click();
    const legendBottom = host.querySelector<HTMLButtonElement>(
      '[data-chart-element="legend-bottom"]',
    );
    expect(legendBottom?.getAttribute('aria-checked')).toBe('true');
    const vertical = host.querySelector<HTMLButtonElement>(
      '[data-chart-element="gridlines-vertical"]',
    );
    expect(vertical?.getAttribute('aria-disabled')).toBe('true');
    const legendRight = host.querySelector<HTMLButtonElement>(
      '[data-chart-element="legend-right"]',
    );
    const pick = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(pick, 'target', { value: legendRight });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(pick)).toBe(true);
    expect(onlyChart().legend).toMatchObject({ visible: true, position: 'right' });

    command('chartColors').click();
    const mono = host.querySelector<HTMLButtonElement>('[data-chart-colors="mono"]');
    const colors = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(colors, 'target', { value: mono });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(colors)).toBe(true);
    expect(chartUsesPalette(onlyChart(), 'mono', workbook)).toBe(true);

    // Each change is its own undo step.
    sheet.instance.history.undo();
    expect(chartUsesPalette(onlyChart(), 'office', workbook)).toBe(true);
    tb.dispose();
  });

  it('opens the chart settings from the Format tab', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    insertChart();
    tb.setActiveTab('chartFormat');
    command('chartFormatSelection').click();
    const settings = document.querySelector<HTMLElement>('.fc-chart-settings');
    expect(settings?.hidden).toBe(false);
    expect(settings?.querySelector('[role="tab"][aria-selected="true"]')?.id).toBe(
      'fc-chart-settings-tab-styles',
    );
    tb.dispose();
  });
});
