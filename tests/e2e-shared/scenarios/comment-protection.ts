import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readThreads = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listThreadedComments()
      .map(({ text, parentId, done }) => ({ text, parentId, done })),
  );

/** On a protected sheet a threaded comment on a locked cell can be added, replied to, edited,
 *  resolved and deleted without a prompt, while a new note stays refused without Edit objects. */
export async function runProtectedSheetCommentScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', platform: 'mac' });
  await sheet.expectNoStub();
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).protect(),
  );
  const alert = page.locator('.fc-valdlg:visible');
  const dialog = page.locator('.fc-cmtnote').first();
  const texts = async () => (await readThreads(page)).map((entry) => entry.text);

  await sheet.goTo('B3');
  await selectRibbonTab(page, 'review');
  await sheet.clickRibbon('mac.review.newComment');
  await expect(dialog).toBeVisible();
  await dialog.locator('textarea').fill('Review this total');
  await dialog.locator('.fc-cmtnote__send').click();
  await expect
    .poll(() => readThreads(page))
    .toMatchObject([{ text: 'Review this total', parentId: null }]);

  await sheet.clickRibbon('mac.review.newComment');
  await dialog.locator('.fc-cmtnote__reply').click();
  await dialog.locator('textarea').fill('Confirmed');
  await dialog.locator('.fc-cmtnote__btn--primary').click();
  await expect.poll(texts).toEqual(['Review this total', 'Confirmed']);

  await sheet.clickRibbon('mac.review.newComment');
  await dialog.locator('.fc-cmtnote__thread-edit').first().click();
  await dialog.locator('textarea').fill('Review this sum');
  await dialog.locator('.fc-cmtnote__btn--primary').click();
  await expect.poll(texts).toEqual(['Review this sum', 'Confirmed']);

  await sheet.clickRibbon('mac.review.newComment');
  await dialog.locator('.fc-cmtnote__resolve').click();
  await page.keyboard.press('Escape');
  await expect.poll(async () => (await readThreads(page))[0]?.done).toBe(true);

  await sheet.clickRibbon('mac.review.deleteComment');
  await expect.poll(() => readThreads(page)).toEqual([]);
  await expect(alert).toHaveCount(0);

  // A note shape is locked, so a new note needs Edit objects.
  await sheet.clickRibbon('mac.review.notes');
  await expect(page.locator('[data-ribbon-command="mac.review.newNote"]').first()).toBeDisabled();
  await page.keyboard.press('Escape');
  await sheet.expectNoConsoleErrors();
}
