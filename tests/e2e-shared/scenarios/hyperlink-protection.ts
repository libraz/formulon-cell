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

/** On a protected sheet, the Insert hyperlinks option lets a link reach an unlocked cell only; a
 *  locked cell keeps refusing it. Without the option the Link command is unavailable. */
export async function runProtectedSheetHyperlinkScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ B2: 'unlocked', B3: 'locked' });
  await selectRibbonTab(page, 'insert');
  const link = page.locator('[data-ribbon-command="hyperlinkInsert"]').first();

  await protectWithUnlockedB2(page, true);
  await sheet.goTo('B2');
  await expect(link).toBeEnabled();
  await linkWithCtrlK(sheet, 'B2', 'https://example.com/unlocked');
  await expect.poll(() => linkAt(sheet, 'B2')).toBe('https://example.com/unlocked');
  await linkWithCtrlK(sheet, 'B3', 'https://example.com/locked');
  await expect.poll(() => linkAt(sheet, 'B3')).toBeNull();

  await protectWithUnlockedB2(page, false);
  await sheet.goTo('B2');
  await expect(link).toBeDisabled();
  await sheet.expectNoConsoleErrors();
}
