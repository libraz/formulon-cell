import { expect, type Page } from '@playwright/test';

import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

/** Runs Review > Spelling on the current selection and returns the report's text, then closes it. */
const spellingReport = async (sheet: UserJourneyPage, page: Page): Promise<string> => {
  await selectRibbonTab(page, 'review');
  await sheet.clickRibbon('spellingReview');
  const dialog = page.getByRole('dialog', { name: 'Spelling', exact: true });
  await expect(dialog).toBeVisible();
  const text = (await dialog.textContent()) ?? '';
  await dialog.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(dialog).toBeHidden();
  return text;
};

/** Spelling checks the whole sheet when one cell is selected and only the selection when a range
 *  is selected, as desktop Excel does. */
export async function runSpellingScopeScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'teh teh' });

  await sheet.goTo('D10');
  expect(await spellingReport(sheet, page)).toContain('Possible typo: "teh"');

  await sheet.goTo('C1:D2');
  const scoped = await spellingReport(sheet, page);
  expect(scoped).toContain('No issues found.');
  expect(scoped).not.toContain('teh');

  await sheet.goTo('A1:B2');
  expect(await spellingReport(sheet, page)).toContain('Possible typo: "teh"');
  await sheet.expectNoConsoleErrors();
}
