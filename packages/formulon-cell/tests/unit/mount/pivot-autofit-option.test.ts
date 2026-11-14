import { PivotAggregation, PivotAxis } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  legacyInstanceOf,
  mount,
  type SpreadsheetInstance,
} from '../../../src/api/spreadsheet/instance.js';
import { workbookHandleOf } from '../../../src/api/workbook.js';
import { History } from '../../../src/commands/history.js';
import { finishPivotReportUpdate } from '../../../src/commands/pivot-table.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { openPivotOptions } from '../../../src/interact/pivot-options-dialog.js';
import { mutators, type SpreadsheetStore } from '../../../src/store/store.js';
import { createHostElement, installDomStubs, uninstallDomStubs } from '../../test-utils/dom.js';

/** Report column D (index 3) is drawn narrower than its content. */
const NARROW = 12;
const REPORT_COL = 3;

describe('PivotTable "autofit column widths on update" option', () => {
  let instance: SpreadsheetInstance;
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let host: HTMLElement;
  let cleanup: () => void;

  beforeEach(async () => {
    installDomStubs();
    ({ host, cleanup } = createHostElement());
    instance = await mount(host);
    wb = workbookHandleOf(instance.workbook);
    store = legacyInstanceOf(instance).store;
  });
  afterEach(() => {
    instance.dispose();
    cleanup();
    uninstallDomStubs();
  });

  const widthOfReportColumn = (): number | undefined =>
    store.getState().layout.colWidths.get(REPORT_COL);

  /** A pivot added through the public facade at D1, with the option set before its fields land. */
  const addFacadePivot = (autofit: boolean) => {
    const facade = instance.workbook;
    const s = facade.sheet(0);
    s.setValues('A1', [
      ['Region', 'Amount'],
      ['East region long label', 10],
      ['West', 20],
    ]);
    const cache = facade.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B3', sheet: s });
    cache.refresh();
    mutators.setColWidth(store, REPORT_COL, NARROW);
    const pivot = s.addPivot({ name: 'P', cacheId: cache, anchor: 'D1' });
    pivot.updateOptions({ autoFormatting: autofit });
    mutators.setColWidth(store, REPORT_COL, NARROW);
    pivot.setFieldAxis('Region', PivotAxis.Row);
    pivot.addDataField({ fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
    return pivot;
  };

  it('widens the report columns of a public addPivot report when the option is on', () => {
    addFacadePivot(true);
    expect(widthOfReportColumn()).toBeGreaterThan(NARROW);
  });

  it('keeps the column widths of a public addPivot report when the option is off', () => {
    addFacadePivot(false);
    expect(widthOfReportColumn()).toBe(NARROW);
  });

  it('keeps the column widths on the dialog completion path when the option is off', () => {
    addFacadePivot(false);
    finishPivotReportUpdate(store, wb);
    expect(widthOfReportColumn()).toBe(NARROW);
    wb.setPivotAutofitColumns(0, 0, true);
    finishPivotReportUpdate(store, wb);
    expect(widthOfReportColumn()).toBeGreaterThan(NARROW);
  });

  it('stops autofit after the Options dialog turns the check box off, and reopens showing it off', async () => {
    addFacadePivot(true);
    const history = new History();
    const ctx = {
      store: store,
      origin: 'ribbon' as const,
      commandId: 'test',
      host: host,
      strings: en,
      wb,
      history,
      onChanged: () => {},
    };
    const box = (): HTMLInputElement => {
      const el = document.querySelector<HTMLInputElement>('.fc-pvoptdlg [data-pvopt-autofit]');
      if (!el) throw new Error('autofit check box missing');
      return el;
    };
    const done = openPivotOptions(ctx, 0, 0);
    expect(box().checked).toBe(true);
    expect(box().disabled).toBe(false);
    box().click();
    document.querySelector<HTMLButtonElement>('.fc-pvoptdlg [data-pvopt-ok]')?.click();
    expect(await done).toBe(true);
    expect(wb.pivotAutofitColumns(0, 0)).toBe(false);

    mutators.setColWidth(store, REPORT_COL, NARROW);
    finishPivotReportUpdate(store, wb);
    expect(widthOfReportColumn()).toBe(NARROW);

    const reopened = openPivotOptions(ctx, 0, 0);
    expect(box().checked).toBe(false);
    document.querySelector<HTMLButtonElement>('.fc-pvoptdlg [data-pvopt-cancel]')?.click();
    expect(await reopened).toBe(false);

    expect(history.undo()).toBe(true);
    expect(wb.pivotAutofitColumns(0, 0)).toBe(true);
    expect(history.redo()).toBe(true);
    expect(wb.pivotAutofitColumns(0, 0)).toBe(false);
  });
});
