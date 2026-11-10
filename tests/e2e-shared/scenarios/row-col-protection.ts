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

const item = (menu: Locator, kind: Kind, choice: 'cells' | 'rows' | 'cols') =>
  menu.locator(`[data-cell-${kind}="${choice}"]`);

/** Opens Insert…/Delete… from the cell context menu and returns the cell dialog. */
const openCellDialog = async (sheet: UserJourneyPage, page: Page, kind: Kind) => {
  await sheet.focusHost();
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  await menu.locator(`[data-fc-action="${kind}Cells"]`).click();
  const dialog = page.locator('.fc-cellshift');
  await expect(dialog).toBeVisible();
  return dialog;
};

/** Asserts the cell dialog offers only the entire-row choice: shifts and entire column disabled. */
const expectEntireRowOnly = async (dialog: Locator, kind: Kind) => {
  const shifts = kind === 'insert' ? ['right', 'down'] : ['left', 'up'];
  for (const value of [...shifts, 'cols'])
    await expect(dialog.locator(`input[value="${value}"]`)).toBeDisabled();
  await expect(dialog.locator('input[value="rows"]')).toBeEnabled();
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

const expectText = (sheet: UserJourneyPage, ref: string, value: string) =>
  expect.poll(() => sheet.readValue(ref)).toEqual({ kind: 'text', value });

/** On a protected sheet a cell shift is refused whatever the options, while whole rows follow only
 *  their own option from every entry point: the Home Insert/Delete menus, the context-menu cell
 *  dialog, and Ctrl+Plus/Minus. Group is refused under any protection, and without options Insert
 *  is unavailable. */
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
  let dialog = await openCellDialog(sheet, page, 'insert');
  await expectEntireRowOnly(dialog, 'insert');
  await dialog.locator('input[value="rows"]').check();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expectText(sheet, 'A4', 'two');

  await sheet.goTo('A4');
  await sheet.shortcut('Shift+Equal');
  dialog = page.locator('.fc-cellshift');
  await expect(dialog).toBeVisible();
  await expectEntireRowOnly(dialog, 'insert');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();

  await selectRow(page, 4);
  await sheet.shortcut('Shift+Equal');
  await expectText(sheet, 'A5', 'two');
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);

  // Delete rows only. Column A now reads one, (blank) x3, two, three. The Windows chrome deletes a
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
  await expectText(sheet, 'A4', 'two');

  await sheet.goTo('A2');
  dialog = await openCellDialog(sheet, page, 'delete');
  await expectEntireRowOnly(dialog, 'delete');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();

  await selectRow(page, 2);
  await sheet.shortcut('Minus');
  await expectText(sheet, 'A3', 'two');
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);

  await selectRow(page, 2);
  await sheet.shortcut('Shift+Equal');
  // A refused shortcut changes nothing; give it time to act before reading.
  await page.waitForTimeout(300);
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await expectText(sheet, 'A3', 'two');

  // No options.
  await protectSheet(page, {});
  await selectRibbonTab(page, 'home');
  await expect(page.locator('[data-ribbon-command="insertRows"]').first()).toBeDisabled();
  await sheet.expectNoConsoleErrors();
}
