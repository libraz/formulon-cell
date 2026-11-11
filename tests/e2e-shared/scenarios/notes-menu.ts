import { expect, type Locator, type Page } from '@playwright/test';

import type { ProtectOptions } from '../../../packages/formulon-cell/src/api/sheet-format.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const protectSheet = (page: Page, options?: ProtectOptions) =>
  page.evaluate(async (opts) => {
    const sheet = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await sheet.unprotect();
    await sheet.protect(opts);
  }, options);

/** Opens Review > Notes at `ref` and returns the menu's first row. */
const openNoteRow = async (sheet: UserJourneyPage, page: Page, ref: string): Promise<Locator> => {
  await sheet.goTo(ref);
  await selectRibbonTab(page, 'review');
  await sheet.clickRibbon('mac.review.notes');
  return page.locator('.fc-tb__menu [data-ribbon-command="mac.review.newNote"]').first();
};

const closeMenu = async (page: Page, row: Locator) => {
  await page.keyboard.press('Escape');
  await expect(row).toBeHidden();
};

/** The Mac Notes menu's first row reads New Note on a plain cell and Edit Note on a note cell.
 *  On a protected sheet without Edit objects the Edit Note row is absent, and New Note stays
 *  greyed; with Edit objects both come back. */
export async function runMacNotesMenuScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', platform: 'mac', locale: 'ja' });
  await sheet.expectNoStub();

  let row = await openNoteRow(sheet, page, 'B2');
  await expect(row).toHaveText('新しいメモ');
  await row.click();
  const dialog = page.locator('.fc-cmtnote').first();
  await expect(dialog).toBeVisible();
  await dialog.locator('textarea').fill('Check the total');
  await dialog.locator('.fc-cmtnote__btn--primary').click();
  await expect(dialog).toBeHidden();

  row = await openNoteRow(sheet, page, 'B2');
  await expect(row).toHaveText('メモの編集');
  await expect(row).toBeVisible();
  await closeMenu(page, row);

  row = await openNoteRow(sheet, page, 'D4');
  await expect(row).toHaveText('新しいメモ');
  await closeMenu(page, row);

  await protectSheet(page);
  row = await openNoteRow(sheet, page, 'B2');
  await expect(row).toBeHidden();
  await page.keyboard.press('Escape');

  row = await openNoteRow(sheet, page, 'D4');
  await expect(row).toHaveText('新しいメモ');
  await expect(row).toBeDisabled();
  await closeMenu(page, row);

  await protectSheet(page, { allowEditObjects: true });
  row = await openNoteRow(sheet, page, 'B2');
  await expect(row).toHaveText('メモの編集');
  await expect(row).toBeEnabled();
  await closeMenu(page, row);
  await sheet.expectNoConsoleErrors();
}
