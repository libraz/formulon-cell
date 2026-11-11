import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Protects the sheet with only the cells in `unlocked` editable. */
const protectWithUnlocked = (page: Page, unlocked: string) =>
  page.evaluate(async (range) => {
    const target = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await target.unprotect();
    target.updateCellFormat('B2:B4', { protection: { locked: true } });
    target.updateCellFormat(range, { protection: { locked: false } });
    await target.protect();
  }, unlocked);

const readColumn = async (sheet: UserJourneyPage) =>
  Promise.all(['B2', 'B3', 'B4'].map((ref) => sheet.readValue(ref)));

const numbers = (...values: number[]) => values.map((value) => ({ kind: 'number', value }));

/** Selects B2:B4, opens the cell context menu over it and returns its Clear Contents item. */
const openClearItem = async (sheet: UserJourneyPage, page: Page) => {
  // Selecting through the name box also hands keyboard focus back to the grid.
  await sheet.goTo('B2:B4');
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  return { menu, item: menu.locator('[data-fc-action="clear"]') };
};

/** On a protected sheet a selection holding any locked cell greys Clear Contents in the Home ribbon
 *  and the context menu, and Delete silently changes nothing, its unlocked cells included. A
 *  selection of unlocked cells only clears. */
export async function runProtectedSheetClearScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ B2: 10, B3: 20, B4: 30 });
  const clear = page.locator('[data-ribbon-command="clearFormat"]').first();
  const alert = page.locator('.fc-valdlg:visible, [role="alertdialog"]:visible');

  await protectWithUnlocked(page, 'B3');
  await sheet.goTo('B2:B4');
  await selectRibbonTab(page, 'home');
  await expect(clear).toBeDisabled();
  let context = await openClearItem(sheet, page);
  await expect(context.item).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(context.menu).toBeHidden();

  await sheet.goTo('B2:B4');
  await page.keyboard.press('Delete');
  // A refused Delete changes nothing; give it time to act before reading.
  await page.waitForTimeout(300);
  await expect(alert).toHaveCount(0);
  expect(await readColumn(sheet)).toEqual(numbers(10, 20, 30));

  await protectWithUnlocked(page, 'B2:B4');
  await sheet.goTo('B2:B4');
  await expect(clear).toBeEnabled();
  context = await openClearItem(sheet, page);
  await expect(context.item).toBeEnabled();
  await context.item.click();
  await expect(context.menu).toBeHidden();
  await expect
    .poll(() => readColumn(sheet))
    .toEqual([{ kind: 'blank' }, { kind: 'blank' }, { kind: 'blank' }]);
  await expect(alert).toHaveCount(0);
  await sheet.expectNoConsoleErrors();
}
