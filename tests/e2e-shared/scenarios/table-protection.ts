import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readTables = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listTables()
      .map((table) => ({
        name: table.name,
        headerRow: table.getProperties().headerRow,
        showRowStripes: table.getStyleOptions()?.showRowStripes ?? null,
      })),
  );

/** On a protected sheet a table keeps only its banding and emphasis options: Banded Rows still
 *  toggles, while Header Row, Insert Table and Ctrl+T are unavailable. */
export async function runProtectedSheetTableScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({
    A1: 'Item',
    B1: 'Qty',
    A2: 'Pen',
    B2: 3,
    A3: 'Ink',
    B3: 5,
    D5: 'Code',
    E5: 'Count',
    D6: 'X1',
    E6: 7,
  });
  await page.evaluate(async () => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    const target = instance.workbook.sheet(0);
    target.addTable({ range: 'A1:B3', name: 'Sales' });
    instance.history.clear();
    await target.protect();
  });
  const before = await readTables(page);
  expect(before).toHaveLength(1);

  await sheet.goTo('A2');
  await selectRibbonTab(page, 'tableDesign');
  const banded = page.locator('[data-ribbon-command="tableBandedRows"]').first();
  await expect(banded).toBeEnabled();
  await banded.click();
  await expect
    .poll(async () => (await readTables(page))[0]?.showRowStripes)
    .toBe(!before[0]?.showRowStripes);
  await expect(page.locator('[data-ribbon-command="tableHeaderRow"]').first()).toBeDisabled();
  expect((await readTables(page))[0]?.headerRow).toBe(true);

  await selectRibbonTab(page, 'insert');
  await expect(page.locator('[data-ribbon-command="formatTableInsert"]').first()).toBeDisabled();

  // D5:E6 is a region Ctrl+T would turn into a table on an unprotected sheet.
  await sheet.goTo('D5');
  await sheet.shortcut('t');
  await page.waitForTimeout(300);
  expect(await readTables(page)).toHaveLength(1);
  await sheet.expectNoConsoleErrors();
}
