import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Protects the sheet under default options with only `unlocked` editable in B2:B5. */
const protectWithUnlocked = (page: Page, unlocked: string) =>
  page.evaluate(async (range) => {
    const target = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await target.unprotect();
    target.updateCellFormat('B2:B5', { protection: { locked: true } });
    target.updateCellFormat(range, { protection: { locked: false } });
    await target.protect();
  }, unlocked);

/** Selects `ref` through the name box, which also hands focus to the grid, and opens the cell menu. */
const openCellMenu = async (sheet: UserJourneyPage, page: Page, ref: string): Promise<Locator> => {
  await sheet.goTo(ref);
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  return menu;
};

const closeMenu = async (page: Page, menu: Locator) => {
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

const expectItems = async (menu: Locator, ids: readonly string[], enabled: boolean) => {
  for (const id of ids) {
    const item = menu.locator(`[data-fc-action="${id}"]`);
    if (enabled) await expect(item, id).toBeEnabled();
    else await expect(item, id).toBeDisabled();
  }
};

/** The Mac cell menu with nothing copied, in desktop order; `|` is a separator. */
const MAC_CELL_MENU = [
  'cut',
  'copy',
  'paste',
  'pasteSpecial',
  '|',
  'insertCells',
  'deleteCells',
  'clear',
  '|',
  'filter',
  'sort',
  '|',
  'newComment',
  'insertComment',
  '|',
  'formatCells',
  'pickFromList',
  'showPhonetic',
  'defineName',
  '|',
  'insertHyperlink',
  'openHyperlink',
];

const readMenuOrder = (menu: Locator) =>
  menu.evaluate((root) =>
    [...root.children].map((el) => {
      const node = el as HTMLElement;
      if (node.classList.contains('fc-ctxmenu__sep')) return '|';
      return node.dataset.fcAction ?? node.dataset.fcSubmenu ?? node.className;
    }),
  );

/** Greyed by their Protect Sheet option whatever the selection. */
const OPTION_GATED = ['formatCells', 'showPhonetic', 'defineName', 'insertHyperlink'] as const;

/** On a protected sheet under default options, Format Cells, Show Phonetic, Define Name and
 *  Hyperlink are greyed whatever the selection, while New Comment stays enabled. New Note greys
 *  too: notes follow Edit objects rather than the cell's lock, as Edit Note and Delete Note do.
 *  Clear Contents is greyed when the selection holds a locked cell, and Pick From Drop-down List
 *  when the active cell is locked; on unlocked cells both work, a pick being an undoable entry. */
export async function runProtectedSheetMenuScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform: 'mac' });
  await sheet.expectNoStub();
  await sheet.seed({ B2: 'East', B3: 'West', B4: 'East' });

  await protectWithUnlocked(page, 'B3');
  let menu = await openCellMenu(sheet, page, 'B2:B4');
  await expectItems(menu, [...OPTION_GATED, 'insertComment', 'clear', 'pickFromList'], false);
  await expectItems(menu, ['newComment'], true);
  await closeMenu(page, menu);

  await protectWithUnlocked(page, 'B2:B5');
  menu = await openCellMenu(sheet, page, 'B2:B4');
  await expectItems(menu, [...OPTION_GATED, 'insertComment'], false);
  await expectItems(menu, ['clear', 'newComment', 'pickFromList'], true);
  expect(await readMenuOrder(menu)).toEqual(MAC_CELL_MENU);
  await closeMenu(page, menu);

  // B5 is unlocked and below the column block: the list offers its distinct entries, sorted.
  menu = await openCellMenu(sheet, page, 'B5');
  await menu.locator('[data-fc-action="pickFromList"]').click();
  const list = page.locator('.fc-validation-list:visible');
  await expect(list).toBeVisible();
  await expect(list.locator('.fc-validation-list__item')).toHaveText(['East', 'West']);
  await list.locator('.fc-validation-list__item', { hasText: 'West' }).click();
  await expect.poll(() => sheet.readValue('B5')).toEqual({ kind: 'text', value: 'West' });
  expect(
    await page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.history.canUndo),
  ).toBe(true);
  // A pick hands focus back to the grid, so Cmd+Z undoes it at once.
  await expect(page.locator('.fc-host').first()).toBeFocused();
  await sheet.shortcut('z');
  await expect.poll(() => sheet.readValue('B5')).toEqual({ kind: 'blank' });
  await sheet.expectNoConsoleErrors();
}

/** Pick From Drop-down List offers the distinct text values of the filled block around the active
 *  cell in its column: header included, stopping at a blank, without the active cell itself or
 *  numbers, folded case-insensitively to the first spelling, sorted ascending. */
export async function runPickListContentScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform: 'mac' });
  await sheet.expectNoStub();
  await sheet.seed({
    B1: 'header',
    B2: 'pear',
    B3: 'apple',
    B4: 10,
    B5: 'Apple',
    B6: 'pear',
    B7: 'banana',
    B9: 'cherry',
  });
  const menu = await openCellMenu(sheet, page, 'B7');
  await menu.locator('[data-fc-action="pickFromList"]').click();
  const list = page.locator('.fc-validation-list:visible');
  await expect(list).toBeVisible();
  await expect(list.locator('.fc-validation-list__item')).toHaveText(['apple', 'header', 'pear']);
  await page.keyboard.press('Escape');
  await expect(list).toBeHidden();
  await sheet.expectNoConsoleErrors();
}
