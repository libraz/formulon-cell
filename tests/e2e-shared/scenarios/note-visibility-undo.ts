import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };
type Platform = 'default' | 'mac';

const noteShown = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).getCommentVisible(1, 1),
  );

const showAll = (page: Page) =>
  page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.options.showAllNotes ?? false);

const canUndo = (page: Page) =>
  page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.history.canUndo);

const toggleFromCellMenu = async (sheet: UserJourneyPage, page: Page) => {
  await sheet.goTo('B2');
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await menu.locator('[data-fc-action="showHideNote"]').click();
  await expect(menu).toBeHidden();
};

/** Runs a Review note command from the Notes menu. */
const runNoteCommand = async (
  sheet: UserJourneyPage,
  page: Page,
  platform: Platform,
  command: 'showHide' | 'showAll',
) => {
  await sheet.goTo('B2');
  await selectRibbonTab(page, 'review');
  if (platform === 'default') {
    await sheet.clickRibbon('notesReview');
    const action = command === 'showHide' ? 'toggle-note' : 'toggle-all-notes';
    await page.locator(`#menu-review-notes [data-comment-action="${action}"]`).click();
    return;
  }
  await sheet.clickRibbon('mac.review.notes');
  const id = command === 'showHide' ? 'mac.review.showHideNote' : 'mac.review.showNotes';
  await page.locator(`.fc-tb__menu [data-ribbon-command="${id}"]`).first().click();
};

/** Asserts the undo stack is empty, then presses Cmd/Ctrl+Z on the grid and expects no change. */
const expectNoUndo = async (
  sheet: UserJourneyPage,
  page: Page,
  state: { shown: boolean; all: boolean },
) => {
  expect(await canUndo(page)).toBe(false);
  await expect(
    page
      .getByRole('toolbar', { name: 'Quick Access Toolbar' })
      .getByRole('button', { name: 'Undo', exact: true }),
  ).toBeDisabled();
  await sheet.goTo('B2');
  await sheet.shortcut('z');
  // An empty undo stack changes nothing; give the shortcut time to act before reading.
  await page.waitForTimeout(300);
  expect(await noteShown(page)).toBe(state.shown);
  expect(await showAll(page)).toBe(state.all);
};

/** Showing or hiding a note, one by one or all at once, is not an undo step: it neither pushes nor
 *  clears the undo stack, so Cmd/Ctrl+Z right after it undoes the previous real edit and leaves the
 *  note as it is. */
export async function runNoteVisibilityUndoScenario(page: Page, platform: Platform): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ fixture: 'empty', platform });
  await sheet.expectNoStub();
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .setComment(1, 1, { text: 'Check the total', author: 'QA' }),
  );
  await sheet.seed({ A5: 'base' });
  expect(await noteShown(page)).toBe(false);

  await toggleFromCellMenu(sheet, page);
  await expect.poll(() => noteShown(page)).toBe(true);
  await expectNoUndo(sheet, page, { shown: true, all: false });

  await runNoteCommand(sheet, page, platform, 'showHide');
  await expect.poll(() => noteShown(page)).toBe(false);
  await expectNoUndo(sheet, page, { shown: false, all: false });

  await runNoteCommand(sheet, page, platform, 'showAll');
  await expect.poll(() => showAll(page)).toBe(true);
  await expectNoUndo(sheet, page, { shown: false, all: true });

  await runNoteCommand(sheet, page, platform, 'showAll');
  await expect.poll(() => showAll(page)).toBe(false);

  // A real edit stays on top of the stack across a note toggle.
  await sheet.enter('A1', 'x');
  await expect.poll(() => sheet.readValue('A1')).toEqual({ kind: 'text', value: 'x' });
  await toggleFromCellMenu(sheet, page);
  await expect.poll(() => noteShown(page)).toBe(true);
  expect(await canUndo(page)).toBe(true);
  await sheet.shortcut('z');
  await expect.poll(() => sheet.readValue('A1')).toEqual({ kind: 'blank' });
  expect(await noteShown(page)).toBe(true);
  expect(await canUndo(page)).toBe(false);
  await sheet.expectNoConsoleErrors();
}
