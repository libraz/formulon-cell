import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createPivotTableAtPlacement } from '../../../../src/commands/pivot-table.js';
import { setProtectedSheet } from '../../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

describe('PivotTable tabs of the ribbon', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    for (const [row, cols] of [
      ['Region', 'Sales'],
      ['East', 1],
      ['West', 2],
      ['East', 3],
    ].entries()) {
      for (const [col, value] of cols.entries()) {
        if (typeof value === 'number') workbook.setNumber({ sheet: 0, row, col }, value);
        else workbook.setText({ sheet: 0, row, col }, value);
      }
    }
    sheet = await mountStubSheet({ locale: 'ja', workbook });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const tabIds = (): string[] =>
    Array.from(host.querySelectorAll<HTMLElement>('[data-ribbon-tab]')).map(
      (tab) => tab.dataset.ribbonTab ?? '',
    );

  it('shows Analyze and Design while the active cell is inside a report', () => {
    const store = sheet.instance.store;
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    expect(tabIds()).not.toContain('pivotAnalyze');
    const result = createPivotTableAtPlacement(
      {
        origin: 'ribbon' as const,
        commandId: 'test',
        store,
        workbook: sheet.workbook,
        history: sheet.instance.history,
      },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    expect(result.ok).toBe(true);
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    expect(tabIds().slice(-2)).toEqual(['pivotAnalyze', 'pivotDesign']);
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    expect(tabIds()).not.toContain('pivotAnalyze');
    tb.dispose();
  });

  it('shows no PivotTable tabs on a protected sheet that does not allow PivotTable use', () => {
    const store = sheet.instance.store;
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    createPivotTableAtPlacement(
      {
        origin: 'ribbon' as const,
        commandId: 'test',
        store,
        workbook: sheet.workbook,
        history: sheet.instance.history,
      },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    expect(tabIds()).toContain('pivotAnalyze');
    setProtectedSheet(store, 0, true);
    expect(tabIds()).not.toContain('pivotAnalyze');
    expect(tabIds()).not.toContain('pivotDesign');
    setProtectedSheet(store, 0, true, { permissions: { pivotTables: true } });
    expect(tabIds().slice(-2)).toEqual(['pivotAnalyze', 'pivotDesign']);
    tb.dispose();
  });

  it('toggles the Design options through the style element, one undo step each', () => {
    const store = sheet.instance.store;
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    createPivotTableAtPlacement(
      {
        origin: 'ribbon' as const,
        commandId: 'test',
        store,
        workbook: sheet.workbook,
        history: sheet.instance.history,
      },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    tb.setActiveTab('pivotDesign');
    const pivot = sheet.workbook.getPivotTables()[0];
    if (!pivot) throw new Error('no pivot');
    const info = () => sheet.workbook.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex);
    const stripes = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotRowStripes"]',
    );
    if (!stripes) throw new Error('missing Banded Rows');
    const before = info()?.showRowStripes ?? false;
    stripes.click();
    expect(info()?.showRowStripes).toBe(!before);
    expect(sheet.instance.history.undo()).toBe(true);
    expect(info()?.showRowStripes ?? false).toBe(before);
    tb.dispose();
  });

  it('toggles +/- Buttons and Field Headers on the report, pressed as set, one undo step each', () => {
    const store = sheet.instance.store;
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    createPivotTableAtPlacement(
      {
        origin: 'ribbon' as const,
        commandId: 'test',
        store,
        workbook: sheet.workbook,
        history: sheet.instance.history,
      },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    tb.setActiveTab('pivotAnalyze');
    const options = () => sheet.workbook.pivotOptions(0, 0);
    for (const [command, key] of [
      ['pivotButtons', 'showDrill'],
      ['pivotHeaders', 'showHeaders'],
    ] as const) {
      const button = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      if (!button) throw new Error(`missing ${command}`);
      expect(button.disabled, command).toBe(false);
      expect(button.getAttribute('aria-pressed'), command).toBe('true');
      button.click();
      expect(options()?.[key], command).toBe(false);
      expect(button.getAttribute('aria-pressed'), command).toBe('false');
      expect(sheet.instance.history.undo()).toBe(true);
      expect(options()?.[key], command).toBe(true);
    }
    tb.dispose();
  });

  it('runs Expand Field and Collapse Field on row labels, greyed where they have no rule', () => {
    const store = sheet.instance.store;
    const wb = sheet.workbook;
    for (const [row, product] of ['Product', 'A', 'B', 'B'].entries()) {
      wb.setText({ sheet: 0, row, col: 2 }, product);
    }
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    createPivotTableAtPlacement(
      { origin: 'ribbon' as const, commandId: 'test', store, workbook: wb, history: null },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 2 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    expect(wb.setPivotFieldAxis(0, 0, 2, 0)).toBe(true);
    // East on row 9 with A and B under it, West on row 12, the grand total on row 14.
    const button = (command: string) => {
      const found = host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
      if (!found) throw new Error(`missing ${command}`);
      return found;
    };
    const states = (row: number, col = 4) => {
      mutators.setActive(store, { sheet: 0, row, col });
      tb.setActiveTab('pivotAnalyze');
      return [button('pivotExpandField').disabled, button('pivotCollapseField').disabled];
    };
    expect(states(9)).toEqual([false, false]);
    // The innermost field's Expand Field opens the Show Detail picker.
    expect(states(10)).toEqual([false, false]);
    // Value, grand-total and caption cells grey both.
    expect(states(9, 5)).toEqual([true, true]);
    expect(states(10, 5)).toEqual([true, true]);
    expect(states(14)).toEqual([true, true]);
    expect(states(14, 5)).toEqual([true, true]);
    expect(states(8)).toEqual([true, true]);
    expect(states(8, 5)).toEqual([true, true]);
    states(10);
    button('pivotExpandField').click();
    const picker = document.querySelector<HTMLElement>('.fc-pivotdetail');
    expect(picker?.getAttribute('aria-label')).toBe('詳細を表示');
    expect(
      Array.from(picker?.querySelectorAll('[role="option"]') ?? []).map((o) => o.textContent),
    ).toEqual(['Sales']);
    Array.from(picker?.querySelectorAll<HTMLButtonElement>('button') ?? [])
      .find((b) => b.textContent === 'キャンセル')
      ?.click();
    expect(document.querySelector('.fc-pivotdetail')).toBeNull();
    const detail = () => wb.getPivotField(0, 0, 0)?.items.map((i) => i.showDetail);
    states(10);
    button('pivotCollapseField').click();
    expect(detail()).toEqual([false, false]);
    expect(sheet.instance.history.undo()).toBe(true);
    expect(detail()).toEqual([true, true]);
    tb.dispose();
  });

  it('refreshes the report from its source through the Analyze tab', () => {
    const store = sheet.instance.store;
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    createPivotTableAtPlacement(
      {
        origin: 'ribbon' as const,
        commandId: 'test',
        store,
        workbook: sheet.workbook,
        history: sheet.instance.history,
      },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    tb.setActiveTab('pivotAnalyze');
    const refresh = host.querySelector<HTMLButtonElement>('[data-ribbon-command="pivotRefresh"]');
    if (!refresh) throw new Error('missing Refresh');
    expect(refresh.disabled).toBe(false);
    sheet.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 100);
    refresh.click();
    const totals = Array.from(sheet.instance.store.getState().data.cells.values()).filter(
      (cell) => cell.value.kind === 'number' && cell.value.value === 105,
    );
    expect(totals.length).toBeGreaterThan(0);
    tb.dispose();
  });

  it('opens the field dialog from the Analyze tab', () => {
    const store = sheet.instance.store;
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    createPivotTableAtPlacement(
      {
        origin: 'ribbon' as const,
        commandId: 'test',
        store,
        workbook: sheet.workbook,
        history: sheet.instance.history,
      },
      { source: { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, rowField: 'Region', valueField: 'Sales' },
      { row: 8, col: 4 },
    );
    mutators.setActive(store, { sheet: 0, row: 9, col: 4 });
    tb.setActiveTab('pivotAnalyze');
    const open = vi.spyOn(sheet.instance, 'openPivotFieldSettings').mockReturnValue(true);
    const button = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotFieldSettings"]',
    );
    if (!button) throw new Error('missing Field Settings');
    expect(button.disabled).toBe(false);
    button.click();
    expect(open).toHaveBeenCalledOnce();
    tb.dispose();
  });
});
