import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readTableRect = (page: Page) =>
  page.evaluate(
    () =>
      (window as unknown as InstanceWindow).__fcInst.workbook
        .sheet(0)
        .table('Sales')
        .getProperties().range.rect,
  );

const rect = (lastRow: number) => ({ firstRow: 0, firstCol: 0, lastRow, lastCol: 1 });

const deleteSheetRows = async (sheet: UserJourneyPage, page: Page, range: string) => {
  await sheet.goTo(range);
  await selectRibbonTab(page, 'home');
  await sheet.deleteRowsOrColumns('rows');
};

/** Deleting table rows from Home > Delete keeps the table in step with the engine: one row shrinks
 *  it by one and undo restores it exactly; deleting every data row keeps the first one, blank, and
 *  undo brings every row back. */
export async function runTableRowDeleteScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Item', B1: 'Qty', A2: 'Pen', B2: 3, A3: 'Ink', B3: 5, A4: 'Pad', B4: 2 });
  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    instance.workbook.sheet(0).addTable({ range: 'A1:B4', name: 'Sales' });
    instance.history.clear();
  });
  expect(await readTableRect(page)).toEqual(rect(3));

  await deleteSheetRows(sheet, page, 'A3');
  await expect.poll(() => readTableRect(page)).toEqual(rect(2));
  await expect.poll(() => sheet.readValue('A3')).toEqual({ kind: 'text', value: 'Pad' });
  await sheet.focusHost();
  await sheet.shortcut('z');
  await expect.poll(() => readTableRect(page)).toEqual(rect(3));
  await expect.poll(() => sheet.readValue('A3')).toEqual({ kind: 'text', value: 'Ink' });

  await deleteSheetRows(sheet, page, 'A2:A4');
  await expect.poll(() => readTableRect(page)).toEqual(rect(1));
  await expect.poll(() => sheet.readValue('A2')).toEqual({ kind: 'blank' });
  await expect.poll(() => sheet.readValue('A1')).toEqual({ kind: 'text', value: 'Item' });

  await sheet.focusHost();
  await sheet.shortcut('z');
  await expect.poll(() => readTableRect(page)).toEqual(rect(3));
  for (const [ref, value] of [
    ['A2', 'Pen'],
    ['A3', 'Ink'],
    ['A4', 'Pad'],
  ] as const) {
    await expect.poll(() => sheet.readValue(ref)).toEqual({ kind: 'text', value });
  }
  await expect.poll(() => sheet.readValue('B4')).toEqual({ kind: 'number', value: 2 });

  await sheet.shortcut('y');
  await expect.poll(() => readTableRect(page)).toEqual(rect(1));
  await expect.poll(() => sheet.readValue('A2')).toEqual({ kind: 'blank' });
  await sheet.expectNoConsoleErrors();
}
