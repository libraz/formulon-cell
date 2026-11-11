import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { revealRibbonCommand, selectRibbonTab } from '../pages/ribbon-tabs.js';
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

const TABLE_STYLE_OPTIONS = [
  'tableHeaderRow',
  'tableTotalRow',
  'tableBandedRows',
  'tableFirstColumn',
  'tableLastColumn',
  'tableBandedColumns',
  'tableFilterButton',
] as const;

/** The Refresh menu on a table without a data connection: [action, label, greyed], in menu order. */
const REFRESH_MENU_JA = [
  ['refresh', '更新', true],
  ['refresh-all', 'すべて更新', false],
  ['refresh-status', 'データ更新の状況', true],
  ['cancel-refresh', '更新のキャンセル', true],
  ['connection-properties', '接続のプロパティ…', true],
] as const;

const readWorkbookState = (page: Page) =>
  page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    const target = instance.workbook.sheet(0);
    return {
      cells: target.listCells(),
      tables: target.listTables().map((table) => table.name),
      canUndo: instance.history.canUndo,
    };
  });

/** The Refresh split button is enabled; its menu greys everything but Refresh All, which runs no
 *  connection and so changes nothing. */
const expectRefreshMenu = async (page: Page) => {
  await selectRibbonTab(page, 'tableDesign');
  const button = page.locator('[data-ribbon-command="tableRefresh"]').first();
  await revealRibbonCommand(button);
  await expect(button).toBeEnabled();
  await expect(button).toHaveAttribute('data-ribbon-activation', 'splitPrimary');
  await button.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator('#menu-table-refresh');
  await expect(menu).toBeVisible();
  const items = menu.locator('[data-table-refresh]');
  await expect(items).toHaveCount(REFRESH_MENU_JA.length);
  for (const [index, [action, label, greyed]] of REFRESH_MENU_JA.entries()) {
    const entry = items.nth(index);
    await expect(entry).toHaveAttribute('data-table-refresh', action);
    await expect(entry).toHaveText(label);
    if (greyed) await expect(entry).toBeDisabled();
    else await expect(entry).toBeEnabled();
  }
  const before = await readWorkbookState(page);
  await items.nth(1).click();
  await expect(menu).toBeHidden();
  expect(await readWorkbookState(page)).toEqual(before);
};

/** On a protected sheet the whole Table Design tab is unavailable, the table name and every style
 *  option included, and so are Insert Table and Ctrl+T. Refresh stays enabled, before and under
 *  protection, with the same menu. */
export async function runProtectedSheetTableScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform });
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
  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    instance.workbook.sheet(0).addTable({ range: 'A1:B3', name: 'Sales' });
    instance.history.clear();
  });
  await sheet.goTo('A2');
  await expectRefreshMenu(page);

  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).protect(),
  );
  const before = await readTables(page);
  expect(before).toHaveLength(1);

  await sheet.goTo('A2');
  await selectRibbonTab(page, 'tableDesign');
  await expect(page.locator('[data-ribbon-input="tableName"] input').first()).toBeDisabled();
  for (const command of TABLE_STYLE_OPTIONS)
    await expect(page.locator(`[data-ribbon-command="${command}"]`).first()).toBeDisabled();
  expect(await readTables(page)).toEqual(before);
  await expectRefreshMenu(page);

  if (platform === 'default') {
    await selectRibbonTab(page, 'insert');
    await expect(page.locator('[data-ribbon-command="formatTableInsert"]').first()).toBeDisabled();
  }

  // D5:E6 is a region Ctrl+T would turn into a table on an unprotected sheet.
  await sheet.goTo('D5');
  await sheet.shortcut('t');
  await page.waitForTimeout(300);
  expect(await readTables(page)).toHaveLength(1);
  await sheet.expectNoConsoleErrors();
}
