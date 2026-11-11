import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Protects sheet 0 with B2 unlocked and the Insert hyperlinks option set as given. */
const protectWithUnlockedB2 = (page: Page, allowInsertHyperlinks: boolean) =>
  page.evaluate(async (allow) => {
    const sheet = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await sheet.unprotect();
    sheet.updateCellFormat('B2', { protection: { locked: false } });
    await sheet.protect({ allowInsertHyperlinks: allow });
  }, allowInsertHyperlinks);

const linkWithCtrlK = async (sheet: UserJourneyPage, ref: string, url: string) => {
  await sheet.goTo(ref);
  const dialog = await sheet.openHyperlinkDialog();
  await dialog.locator('input[type="url"]').fill(url);
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
};

const linkAt = async (sheet: UserJourneyPage, ref: string) =>
  (await sheet.readFormat(ref))?.hyperlink ?? null;

/** Opens Home > Clear at `ref` and returns its Remove Hyperlinks row. */
const openRemoveHyperlinks = async (sheet: UserJourneyPage, page: Page, ref: string) => {
  await sheet.goTo(ref);
  await selectRibbonTab(page, 'home');
  await sheet.clickRibbon('clearFormat');
  const menu = page.locator('#menu-clear');
  await expect(menu).toBeVisible();
  return menu.getByRole('menuitem', { name: 'Remove Hyperlinks' });
};

/** On a protected sheet, the Insert hyperlinks option lets a link reach or leave an unlocked cell
 *  only; a locked cell keeps refusing both. Without the option the Link command, Ctrl+K and Remove
 *  Hyperlinks are unavailable even on an unlocked cell. */
export async function runProtectedSheetHyperlinkScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ B2: 'unlocked', B3: 'locked' });
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .setHyperlink('B3', { target: 'https://example.com/original' }),
  );
  await selectRibbonTab(page, 'insert');
  const link = page.locator('[data-ribbon-command="hyperlinkInsert"]').first();

  await protectWithUnlockedB2(page, true);
  await sheet.goTo('B2');
  await expect(link).toBeEnabled();
  await linkWithCtrlK(sheet, 'B2', 'https://example.com/unlocked');
  await expect.poll(() => linkAt(sheet, 'B2')).toBe('https://example.com/unlocked');
  await linkWithCtrlK(sheet, 'B3', 'https://example.com/locked');
  await expect.poll(() => linkAt(sheet, 'B3')).toBe('https://example.com/original');

  await protectWithUnlockedB2(page, false);
  await sheet.goTo('B2');
  await expect(link).toBeDisabled();
  await sheet.shortcut('k');
  await expect(page.locator('.fc-hldlg')).toBeHidden();

  let remove = await openRemoveHyperlinks(sheet, page, 'B2');
  await expect(remove).toBeDisabled();
  await page.keyboard.press('Escape');

  await protectWithUnlockedB2(page, true);
  remove = await openRemoveHyperlinks(sheet, page, 'B2');
  await remove.click();
  await expect.poll(() => linkAt(sheet, 'B2')).toBeNull();
  // Over a locked cell the whole Clear menu is greyed, so its link stays.
  await sheet.goTo('B3');
  await expect(page.locator('[data-ribbon-command="clearFormat"]').first()).toBeDisabled();
  expect(await linkAt(sheet, 'B3')).toBe('https://example.com/original');
  await sheet.expectNoConsoleErrors();
}
