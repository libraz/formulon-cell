import { expect, type Locator, type Page } from '@playwright/test';
import type { ProtectOptions } from '../../../packages/formulon-cell/src/api/sheet-format.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };
type Kind = 'insert' | 'delete';

const protectSheet = (page: Page, options: ProtectOptions) =>
  page.evaluate(async (opts) => {
    const sheet = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await sheet.unprotect();
    await sheet.protect(opts);
  }, options);

const openHomeMenu = async (sheet: UserJourneyPage, page: Page, kind: Kind): Promise<Locator> => {
  await selectRibbonTab(page, 'home');
  await sheet.clickRibbon(kind === 'insert' ? 'insertRows' : 'deleteRows');
  const menu = page
    .locator(kind === 'insert' ? '#menu-insert-cells' : '#menu-delete-cells')
    .first();
  await expect(menu).toBeVisible();
  return menu;
};

const expectText = (sheet: UserJourneyPage, ref: string, value: string) =>
  expect.poll(() => sheet.readValue(ref)).toEqual({ kind: 'text', value });

const item = (menu: Locator, kind: Kind, choice: 'cells' | 'rows' | 'cols') =>
  menu.locator(`[data-cell-${kind}="${choice}"]`);

/** Asserts the cell context menu greys Insert…/Delete…, then closes the menu. */
const expectCellMenuItemGreyed = async (sheet: UserJourneyPage, page: Page, kind: Kind) => {
  await sheet.focusHost();
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  await expect(menu.locator(`[data-fc-action="${kind}Cells"]`)).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

/** Presses Ctrl/Cmd+Plus or Minus and asserts it neither opens the cell dialog nor edits `ref`. */
const expectShortcutNoOp = async (
  sheet: UserJourneyPage,
  page: Page,
  key: string,
  ref: string,
  value: string,
) => {
  await sheet.shortcut(key);
  // A refused shortcut changes nothing; give it time to act before reading.
  await page.waitForTimeout(300);
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await expectText(sheet, ref, value);
};

/** Selects one whole row through the name box and waits until the selection spans it. */
const selectRow = async (page: Page, row: number) => {
  const nameBox = page.locator('.fc-host__formulabar-tag').first();
  await nameBox.fill(`${row}:${row}`);
  await nameBox.press('Enter');
  await expect
    .poll(() =>
      page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.selection.get().range),
    )
    .toMatchObject({ firstRow: row - 1, lastRow: row - 1, firstCol: 0, lastCol: 16383 });
};

/** On a protected sheet a cell shift is refused whatever the options: with a cell selected the
 *  context menu greys Insert…/Delete… and Ctrl+Plus/Minus do nothing. Whole rows follow only their
 *  own option, from the Home Insert/Delete menus and from the shortcuts on a selected row. Group is
 *  refused under any protection, and without options Insert is unavailable. */
export async function runProtectedSheetRowColumnScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'one', A2: 'two', A3: 'three' });

  // Insert rows only.
  await protectSheet(page, { allowInsertRows: true });
  await sheet.goTo('A2');
  let menu = await openHomeMenu(sheet, page, 'insert');
  await expect(item(menu, 'insert', 'cells')).toBeDisabled();
  await expect(item(menu, 'insert', 'cols')).toBeDisabled();
  await item(menu, 'insert', 'rows').click();
  await expectText(sheet, 'A3', 'two');

  await selectRibbonTab(page, 'data');
  await expect(page.locator('[data-ribbon-command="outlineGroup"]').first()).toBeDisabled();

  await sheet.goTo('A3');
  await expectCellMenuItemGreyed(sheet, page, 'insert');
  await expectShortcutNoOp(sheet, page, 'Shift+Equal', 'A3', 'two');

  await selectRow(page, 3);
  await sheet.shortcut('Shift+Equal');
  await expectText(sheet, 'A4', 'two');
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);

  // Delete rows only. Column A now reads one, (blank) x2, two, three. The Windows chrome deletes a
  // whole row on a protected sheet only when every cell in it is unlocked.
  await page.evaluate(async () => {
    const target = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await target.unprotect();
    target.updateCellFormat('A1:XFD20', { protection: { locked: false } });
  });
  await protectSheet(page, { allowDeleteRows: true });
  await sheet.goTo('A2');
  menu = await openHomeMenu(sheet, page, 'delete');
  await expect(item(menu, 'delete', 'cells')).toBeDisabled();
  await expect(item(menu, 'delete', 'cols')).toBeDisabled();
  await item(menu, 'delete', 'rows').click();
  await expectText(sheet, 'A3', 'two');

  await sheet.goTo('A2');
  await expectCellMenuItemGreyed(sheet, page, 'delete');
  await expectShortcutNoOp(sheet, page, 'Minus', 'A3', 'two');

  await selectRow(page, 2);
  await sheet.shortcut('Minus');
  await expectText(sheet, 'A2', 'two');
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);

  // Insert rows is off now, so the whole-row shortcut does nothing.
  await selectRow(page, 2);
  await expectShortcutNoOp(sheet, page, 'Shift+Equal', 'A2', 'two');

  // No options.
  await protectSheet(page, {});
  await selectRibbonTab(page, 'home');
  await expect(page.locator('[data-ribbon-command="insertRows"]').first()).toBeDisabled();
  await sheet.expectNoConsoleErrors();
}
