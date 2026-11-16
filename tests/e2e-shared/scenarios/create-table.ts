import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readTableRects = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listTables()
      .map((table) => {
        const { range, headerRow, columns } = table.getProperties();
        return { rect: range.rect, headerRow, columns: columns.map((column) => column.name) };
      }),
  );

const REGION = { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 };

/** Seeds A1:B3 with a header row and puts the active cell inside it. */
const seedRegion = async (sheet: UserJourneyPage) => {
  await sheet.seed({ A1: 'Item', B1: 'Qty', A2: 'Pen', B2: 3, A3: 'Ink', B3: 5 });
  await sheet.goTo('A2');
};

/** Asserts the Create Table dialog proposes the current region with headers and returns it. */
const expectCreateTableDialog = async (page: Page, range: string): Promise<Locator> => {
  const dialog = page.getByRole('dialog', { name: 'Create Table' });
  await expect(dialog).toBeVisible();
  await expect(dialog.locator('input').first()).toHaveValue(range);
  await expect(dialog.getByRole('checkbox', { name: 'My table has headers' })).toBeChecked();
  return dialog;
};

/** Ctrl+T, Ctrl+L and Insert > Table from one cell of a data region open Create Table on the
 *  region with headers; Cancel creates nothing and OK creates the table over the region. */
export async function runCreateTableEntryPointsScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);

  await sheet.shortcut('t');
  let dialog = await expectCreateTableDialog(page, 'Sheet1!$A$1:$B$3');
  expect(await readTableRects(page)).toEqual([]);
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  expect(await readTableRects(page)).toEqual([]);

  await selectRibbonTab(page, 'insert');
  await sheet.goTo('A2');
  await sheet.clickRibbon('formatTableInsert');
  dialog = await expectCreateTableDialog(page, 'Sheet1!$A$1:$B$3');
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
  expect(await readTableRects(page)).toEqual([]);

  await sheet.goTo('A2');
  await sheet.shortcut('l');
  dialog = await expectCreateTableDialog(page, 'Sheet1!$A$1:$B$3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(dialog).toBeHidden();
  await expect
    .poll(() => readTableRects(page))
    .toEqual([{ rect: REGION, headerRow: true, columns: ['Item', 'Qty'] }]);
  await sheet.expectNoConsoleErrors();
}

/** Format as Table opens the style gallery first, then Create Table on the current region. */
export async function runFormatAsTableGalleryScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);

  await selectRibbonTab(page, 'home');
  await sheet.clickRibbon('formatTableHome');
  const swatch = page.locator('.fc-tb__tablestyle-swatch:visible').first();
  await expect(swatch).toBeVisible();
  expect(await readTableRects(page)).toEqual([]);
  await swatch.click();
  const dialog = await expectCreateTableDialog(page, 'Sheet1!$A$1:$B$3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readTableRects(page))
    .toEqual([{ rect: REGION, headerRow: true, columns: ['Item', 'Qty'] }]);
  await sheet.expectNoConsoleErrors();
}

/** The Mac Insert > Table dialog writes the region as a formula reference on the active sheet. */
export async function runMacCreateTableScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await seedRegion(sheet);

  await selectRibbonTab(page, 'insert');
  await sheet.clickRibbon('mac.insert.table');
  const dialog = await expectCreateTableDialog(page, '=$A$1:$B$3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readTableRects(page))
    .toEqual([{ rect: REGION, headerRow: true, columns: ['Item', 'Qty'] }]);
  await sheet.expectNoConsoleErrors();
}

/** Review > Check Accessibility carries Format as Table, which opens Create Table on the current
 *  region, and Unmerge Cells, which unmerges the selection as one undo step. */
export async function runMacAccessibilityMenuScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform: 'mac' });
  await sheet.expectNoStub();
  await seedRegion(sheet);
  await page.evaluate(() => {
    const inst = (window as unknown as InstanceWindow).__fcInst;
    inst.workbook.sheet(0).addMerge('D2:E3');
    inst.history.clear();
  });
  const menu = page.locator('#menu-mac-review-accessibility');
  const item = (id: string) =>
    menu.locator(`[data-ribbon-command="mac.review.accessibility.${id}"]`);
  await selectRibbonTab(page, 'review');

  await sheet.goTo('A2');
  await sheet.clickRibbon('mac.review.accessibility');
  await expect(menu).toBeVisible();
  await expect(item('formatTable')).toBeEnabled();
  await expect(item('unmerge')).toBeEnabled();
  await item('formatTable').click();
  await expect(menu).toBeHidden();
  const dialog = await expectCreateTableDialog(page, '=$A$1:$B$3');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect
    .poll(() => readTableRects(page))
    .toEqual([{ rect: REGION, headerRow: true, columns: ['Item', 'Qty'] }]);

  await expect.poll(() => sheet.readMerge('D2')).toEqual({ sheet: 0, r0: 1, c0: 3, r1: 2, c1: 4 });
  await sheet.goTo('D2');
  await sheet.clickRibbon('mac.review.accessibility');
  await item('unmerge').click();
  await expect.poll(() => sheet.readMerge('D2')).toBeNull();
  await page.locator('.fc-host__grid').focus();
  await sheet.shortcut('z');
  await expect.poll(() => sheet.readMerge('D2')).toEqual({ sheet: 0, r0: 1, c0: 3, r1: 2, c1: 4 });
  await sheet.expectNoConsoleErrors();
}
