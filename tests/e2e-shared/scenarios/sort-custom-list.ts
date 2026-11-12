import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Data > Sort, Order > Custom List...: a list typed one entry per line becomes the level's order
 *  and the rows follow it. */
export async function runSortCustomListScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Level', A2: 'High', A3: 'Low', A4: 'Mid', A5: 'Low' });
  await page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.selection.set('A1:A5'));
  await selectRibbonTab(page, 'data');
  await sheet.clickRibbon(platform === 'mac' ? 'mac.data.sortCustom' : 'sortData');
  const dialog = page.getByRole('dialog').filter({ has: page.locator('.fc-sortdlg__level') });
  await expect(dialog).toBeVisible();

  const order = dialog
    .locator('.fc-sortdlg__level')
    .first()
    .locator('select.fc-sortdlg__direction');
  await order.selectOption('custom');
  const lists = page.locator('.fc-customlists');
  await expect(lists).toBeVisible();
  await lists.locator('.fc-customlists__entries').fill('Low\nMid\nHigh');
  await lists.locator('.fc-fmtdlg__btn--primary').click();
  await expect(lists).toBeHidden();
  await expect(order).toHaveValue('list');
  await expect(order.locator('option:checked')).toHaveText('Low, Mid, High');
  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();

  const column = async () =>
    (await Promise.all(['A2', 'A3', 'A4', 'A5'].map((ref) => sheet.readValue(ref)))).map((cell) =>
      cell.kind === 'text' ? cell.value : null,
    );
  await expect.poll(column).toEqual(['Low', 'Low', 'Mid', 'High']);
  await sheet.expectNoConsoleErrors();
}
