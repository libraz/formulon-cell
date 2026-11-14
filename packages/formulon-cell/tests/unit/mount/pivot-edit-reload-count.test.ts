import '../../api/setup-dom.js';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { workbookHandleOf, workbookOf } from '../../../src/api/workbook.js';
import {
  createPivotTableFromRange,
  executeRibbonPivotTableAction,
} from '../../../src/commands/pivot-table.js';
import { PivotAggregation, PivotReportLayout } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mount } from '../../../src/index.js';
import { spreadsheetInternals } from '../../../src/mount/internals.js';
import { mutators } from '../../../src/store/store.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';

const setup = async () => {
  const { host, cleanup } = createHostElement();
  const wb = await WorkbookHandle.createDefault();
  const rows: (string | number)[][] = [
    ['Kind', 'Amt'],
    ['a', 1],
    ['b', 2],
  ];
  for (const [r, row] of rows.entries()) {
    for (const [c, v] of row.entries()) {
      if (typeof v === 'number') wb.setNumber({ sheet: 0, row: r, col: c }, v);
      else wb.setText({ sheet: 0, row: r, col: c }, v);
    }
  }
  expect(
    createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      destination: { sheet: 0, row: 0, col: 3 },
      name: 'PivotTable1',
      rowField: 'Kind',
      valueField: 'Amt',
      aggregation: PivotAggregation.Sum,
    }).ok,
  ).toBe(true);
  const instance = await mount(host, { workbook: workbookOf(wb, true), locale: 'en' });
  return { host, cleanup, instance, handle: workbookHandleOf(instance.workbook) };
};

describe('store reloads per pivot action', () => {
  beforeEach(() => installDomStubs());
  afterEach(() => {
    vi.restoreAllMocks();
    uninstallDomStubs();
  });

  const spies = () => {
    const replace = vi.spyOn(mutators, 'replaceCells');
    const refresh = vi.spyOn(mutators, 'refreshActiveSheetCells');
    return () => replace.mock.calls.length + refresh.mock.calls.length;
  };

  it('source edit then Refresh', async () => {
    const { instance, handle, cleanup } = await setup();
    instance.selection.set('D2');
    handle.setNumber({ sheet: 0, row: 1, col: 1 }, 5);
    await new Promise((r) => setTimeout(r, 0));
    const count = spies();
    executeRibbonPivotTableAction({
      store: spreadsheetInternals(instance).store,
      workbook: handle,
      origin: 'ribbon',
      commandId: 'pivotRefresh',
      action: 'refresh',
      history: null,
      strings: {} as never,
    });
    expect(count()).toBe(1);
    cleanup();
  });

  it('facade pivot change', async () => {
    const { instance, cleanup } = await setup();
    const count = spies();
    instance.workbook.sheet(0).pivot(0).setReportLayout(PivotReportLayout.Tabular);
    expect(count()).toBe(1);
    cleanup();
  });

  it('Options dialog OK', async () => {
    const { instance, handle, cleanup } = await setup();
    const count = spies();
    instance.dialogs.open('pivotOptions', { sheet: 0, pivotIndex: 0 });
    const ok = document.querySelector<HTMLButtonElement>('[data-pvopt-ok]');
    const name = document.querySelector<HTMLInputElement>('[data-pvopt-name]');
    if (name) name.value = 'Renamed';
    ok?.click();
    await new Promise((r) => setTimeout(r, 0));
    expect(ok).not.toBeNull();
    expect(handle.pivotTableName(0, 0)).toBe('Renamed');
    expect(count()).toBe(1);
    cleanup();
  });
});
