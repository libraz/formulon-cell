import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const spellingDialog = (page: Page) => page.getByRole('dialog', { name: /^Spelling/ });

const runSpelling = async (sheet: UserJourneyPage, page: Page) => {
  await selectRibbonTab(page, 'review');
  await sheet.clickRibbon('spellingReview');
};

/** Spelling checks the whole sheet when one cell is selected and only the selection when a range
 *  is selected, as desktop Excel does. 'teh teh' flags "teh" as not in the dictionary, then the
 *  second "teh" as a repeated word, and the dialog closes silently after the last one. */
export async function runSpellingScopeScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'teh teh' });
  const dialog = spellingDialog(page);
  const word = dialog.locator('#fc-spelling-word');

  await sheet.goTo('D10');
  await runSpelling(sheet, page);
  await expect(dialog).toBeVisible();
  await expect(word).toHaveValue('teh');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();

  // Nothing to flag inside the range, so the dialog does not open.
  await sheet.goTo('C1:D2');
  await runSpelling(sheet, page);
  await page.waitForTimeout(300);
  await expect(dialog).toHaveCount(0);

  await sheet.goTo('A1:B2');
  await runSpelling(sheet, page);
  await expect(dialog).toBeVisible();
  await expect(dialog).toContainText('Not in Dictionary:');
  await expect(word).toHaveValue('teh');
  await dialog.getByRole('button', { name: 'Ignore Once', exact: true }).click();
  await expect(dialog).toContainText('Repeated Word');
  await expect(word).toHaveValue('teh');
  await expect(dialog.locator('#fc-spelling-change-to')).toHaveValue('');
  await dialog.getByRole('button', { name: 'Ignore Once', exact: true }).click();
  await expect(dialog).toBeHidden();
  // The checked range stays selected.
  expect(
    await page.evaluate(() =>
      [...(window as unknown as InstanceWindow).__fcInst.selection.get().ranges].map((r) => [
        r.firstRow,
        r.firstCol,
        r.lastRow,
        r.lastCol,
      ]),
    ),
  ).toEqual([[0, 0, 1, 1]]);
  expect(await sheet.readValue('A1')).toEqual({ kind: 'text', value: 'teh teh' });
  await sheet.expectNoConsoleErrors();
}
