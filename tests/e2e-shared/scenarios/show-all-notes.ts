import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type UpdateEvent = { keys: string[]; source: string };
type InstanceWindow = Window & { __fcInst: SpreadsheetInstance; __fcUpdates: UpdateEvent[] };
type Platform = 'default' | 'mac';

/** Shown state of the notes on B2 and D4, as the workbook stores it. */
const readShown = (page: Page) =>
  page.evaluate(() => {
    const sheet = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    return [sheet.getCommentVisible(1, 1), sheet.getCommentVisible(3, 3)];
  });

const readOption = (page: Page) =>
  page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.options.showAllNotes ?? false);

const saveHex = (page: Page) =>
  page.evaluate(() =>
    Array.from((window as unknown as InstanceWindow).__fcInst.workbook.save().bytes, (b) =>
      b.toString(16).padStart(2, '0'),
    ).join(''),
  );

/** Clicks Show All Notes and returns the button that carries its pressed state. */
const toggleShowAll = async (
  sheet: UserJourneyPage,
  page: Page,
  platform: Platform,
): Promise<Locator> => {
  await selectRibbonTab(page, 'review');
  if (platform === 'default') {
    await sheet.clickRibbon('showAllNotesReview');
    return page.locator('[data-ribbon-command="showAllNotesReview"]').first();
  }
  await sheet.clickRibbon('mac.review.notes');
  const item = page.locator('.fc-tb__menu [data-ribbon-command="mac.review.showNotes"]').first();
  await item.click();
  return item;
};

/** Reopens the Mac Notes menu to read the pressed state, then closes it. */
const expectPressed = async (page: Page, platform: Platform, button: Locator, pressed: boolean) => {
  if (platform === 'mac')
    await page.locator('[data-ribbon-command="mac.review.notes"]').first().click();
  await expect(button).toHaveAttribute('aria-pressed', String(pressed));
  if (platform === 'mac') await page.keyboard.press('Escape');
};

/** Show All Notes is a view toggle: it draws every note without writing any note's own shown
 *  state, its pressed state follows the toggle rather than whether every note is shown, the saved
 *  file is unchanged by it, and the UI toggle reports an update with source 'user'. */
export async function runShowAllNotesScenario(page: Page, platform: Platform): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const w = window as unknown as InstanceWindow;
    const target = w.__fcInst.workbook.sheet(0);
    target.setComment(1, 1, { text: 'Hidden note', author: 'QA' });
    target.setComment(3, 3, { text: 'Shown note', author: 'QA' });
    target.setCommentVisible(3, 3, true);
    w.__fcUpdates = [];
    w.__fcInst.on('update', (event) => {
      w.__fcUpdates.push({ keys: [...event.keys], source: event.source });
    });
  });
  await sheet.goTo('A1');
  expect(await readShown(page)).toEqual([false, true]);
  const before = await saveHex(page);

  const button = await toggleShowAll(sheet, page, platform);
  await expect.poll(() => readOption(page)).toBe(true);
  await expectPressed(page, platform, button, true);
  expect(await readShown(page)).toEqual([false, true]);
  expect(await saveHex(page)).toBe(before);
  expect(
    await page.evaluate(() => (window as unknown as InstanceWindow).__fcUpdates),
  ).toContainEqual({ keys: ['showAllNotes'], source: 'user' });

  await toggleShowAll(sheet, page, platform);
  await expect.poll(() => readOption(page)).toBe(false);
  await expectPressed(page, platform, button, false);
  expect(await readShown(page)).toEqual([false, true]);

  // Every note shown one by one still leaves the toggle off.
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).setCommentVisible(1, 1, true),
  );
  await expectPressed(page, platform, button, false);
  await sheet.expectNoConsoleErrors();
}
