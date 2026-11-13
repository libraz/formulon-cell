import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const openCellMenu = async (page: Page, row: number, col: number) => {
  const point = await gridCellCenter(page, row, col);
  await page.mouse.click(point.x, point.y, { button: 'right' });
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)').first();
  await expect(menu).toBeVisible();
  return menu;
};

/** On the default profile a plain cell's menu offers New Note alone, a cell with a note offers
 *  Edit Note, Delete Note and Show/Hide Note, and the Review tab's Notes > New Note reads Edit
 *  Note while the active cell has one. */
export async function runDefaultNoteMenuScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .setComment(1, 1, { text: 'Check', author: 'QA' }),
  );

  let menu = await openCellMenu(page, 0, 0);
  await expect(menu.locator('[data-fc-action="insertComment"]')).toContainText('New Note');
  await expect(menu.locator('[data-fc-action="deleteComment"]')).toHaveCount(0);
  await expect(menu.locator('[data-fc-action="showHideNote"]')).toHaveCount(0);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();

  menu = await openCellMenu(page, 1, 1);
  await expect(menu.locator('[data-fc-action="insertComment"]')).toContainText('Edit Note');
  await expect(menu.locator('[data-fc-action="deleteComment"]')).toContainText('Delete Note');
  await expect(menu.locator('[data-fc-action="showHideNote"]')).toContainText('Show/Hide Note');
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();

  await selectRibbonTab(page, 'review');
  const newNote = page.locator('#menu-review-notes [data-comment-action="new-note"]');
  for (const [cell, label] of [
    ['A1', 'New Note'],
    ['B2', 'Edit Note'],
  ] as const) {
    await sheet.goTo(cell);
    await sheet.clickRibbon('notesReview');
    await expect(newNote).toHaveText(label);
    await page.keyboard.press('Escape');
    await expect(newNote).toBeHidden();
  }
  await sheet.expectNoConsoleErrors();
}
