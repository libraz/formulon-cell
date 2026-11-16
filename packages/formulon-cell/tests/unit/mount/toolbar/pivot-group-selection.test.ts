import { PivotDatePeriod } from '@libraz/formulon';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createPivotTableFromRange } from '../../../../src/commands/pivot-table.js';
import { PivotAxis } from '../../../../src/engine/types.js';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { loadFixture } from '../../engine/pivot-excel-parity/fixture.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

// Sheet5 shows Date on rows 4-8 under the header on row 3, the grand total on row 9.
describe('PivotTable Analyze › Group Selection', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;
  let index: number;

  beforeEach(async () => {
    const workbook = await loadFixture('pivot_value_date_filters.xlsx');
    index = Array.from({ length: workbook.sheetCount }, (_, i) => i).find(
      (i) => workbook.sheetName(i) === 'Sheet5',
    ) as number;
    sheet = await mountStubSheet({ locale: 'ja', workbook });
    mutators.setSheetIndex(sheet.instance.store, index);
    mutators.replaceCells(sheet.instance.store, workbook.cells(index));
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('opens the Grouping dialog on a date label and is greyed elsewhere', () => {
    const store = sheet.instance.store;
    const wb = sheet.workbook;
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const button = () => {
      const found = host.querySelector<HTMLButtonElement>(
        '[data-ribbon-command="pivotGroupSelection"]',
      );
      if (!found) throw new Error('missing Group Selection');
      return found;
    };
    const disabledAt = (row: number, col = 0) => {
      mutators.setActive(store, { sheet: index, row, col });
      tb.setActiveTab('pivotAnalyze');
      return button().disabled;
    };
    expect(disabledAt(4)).toBe(false);
    expect(disabledAt(4, 1)).toBe(true);
    expect(disabledAt(8)).toBe(true);
    expect(disabledAt(2)).toBe(true);
    disabledAt(4);
    button().click();
    const dialog = document.querySelector<HTMLElement>('.fc-pivotgroup');
    expect(dialog).not.toBeNull();
    dialog?.querySelector<HTMLButtonElement>('.fc-pivotgroup__footer button:last-child')?.click();
    expect(wb.getPivotDefinition(index, 0)?.fields[2]?.dateGroup?.periods).toBe(
      PivotDatePeriod.Months,
    );
    expect(sheet.instance.history.undo()).toBe(true);
    expect(wb.getPivotDefinition(index, 0)?.fields[2]?.dateGroup).toBeNull();
    // A text field has no grouping this release can make.
    expect(wb.setPivotFieldAxis(index, 0, 0, PivotAxis.Row)).toBe(true);
    const region = Array.from(wb.pivotCells(index)).find(
      (c) => c.kind === 1 && c.fieldName === 'Region',
    );
    expect(region && disabledAt(region.addr.row, region.addr.col)).toBe(true);
    tb.dispose();
  });
});

describe('PivotTable Analyze › Group Selection on a number field', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    workbook.setText({ sheet: 0, row: 0, col: 10 }, 'Amt');
    workbook.setText({ sheet: 0, row: 0, col: 11 }, 'Qty');
    for (let i = 1; i <= 12; i += 1) {
      workbook.setNumber({ sheet: 0, row: i, col: 10 }, i * 10);
      workbook.setNumber({ sheet: 0, row: i, col: 11 }, i);
    }
    expect(
      createPivotTableFromRange(workbook, {
        source: { sheet: 0, r0: 0, c0: 10, r1: 12, c1: 11 },
        destination: { sheet: 0, row: 2, col: 0 },
        rowField: 'Amt',
        valueField: 'Qty',
      }).ok,
    ).toBe(true);
    sheet = await mountStubSheet({ locale: 'ja', workbook });
    mutators.replaceCells(sheet.instance.store, workbook.cells(0));
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('opens the number Grouping dialog on a number label', () => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    const label = Array.from(sheet.workbook.pivotCells(0)).find((c) => c.kind === 1);
    if (!label) throw new Error('no label');
    mutators.setActive(sheet.instance.store, label.addr);
    tb.setActiveTab('pivotAnalyze');
    const button = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotGroupSelection"]',
    );
    expect(button?.disabled).toBe(false);
    button?.click();
    const dialog = document.querySelector<HTMLElement>('.fc-pivotgroup');
    dialog?.querySelector<HTMLButtonElement>('.fc-pivotgroup__footer button:last-child')?.click();
    expect(sheet.workbook.getPivotDefinition(0, 0)?.fields[0]?.numericGroup?.interval).toBe(10);
    tb.dispose();
  });
});
