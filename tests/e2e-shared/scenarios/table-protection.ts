import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { revealRibbonCommand, selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readTables = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .sheet(0)
      .listTables()
      .map((table) => ({
        name: table.name,
        headerRow: table.getProperties().headerRow,
        showRowStripes: table.getStyleOptions()?.showRowStripes ?? null,
      })),
  );

const TABLE_STYLE_OPTIONS = [
  'tableHeaderRow',
  'tableTotalRow',
  'tableBandedRows',
  'tableFirstColumn',
  'tableLastColumn',
  'tableBandedColumns',
  'tableFilterButton',
] as const;

/** The Refresh menu on a table without a data connection: [action, label, greyed], in menu order. */
const REFRESH_MENU_JA = [
  ['refresh', '更新', true],
  ['refresh-all', 'すべて更新', false],
  ['refresh-status', 'データ更新の状況', true],
  ['cancel-refresh', '更新のキャンセル', true],
  ['connection-properties', '接続のプロパティ…', true],
] as const;

const readWorkbookState = (page: Page) =>
  page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    const target = instance.workbook.sheet(0);
    return {
      cells: target.listCells(),
      tables: target.listTables().map((table) => table.name),
      canUndo: instance.history.canUndo,
    };
  });

/** The Refresh split button is enabled; its menu greys everything but Refresh All, which runs no
 *  connection and so changes nothing. */
const expectRefreshMenu = async (page: Page) => {
  await selectRibbonTab(page, 'tableDesign');
  const button = page.locator('[data-ribbon-command="tableRefresh"]').first();
  await revealRibbonCommand(button);
  await expect(button).toBeEnabled();
  await expect(button).toHaveAttribute('data-ribbon-activation', 'splitPrimary');
  await button.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator('#menu-table-refresh');
  await expect(menu).toBeVisible();
  const items = menu.locator('[data-table-refresh]');
  await expect(items).toHaveCount(REFRESH_MENU_JA.length);
  for (const [index, [action, label, greyed]] of REFRESH_MENU_JA.entries()) {
    const entry = items.nth(index);
    await expect(entry).toHaveAttribute('data-table-refresh', action);
    await expect(entry).toHaveText(label);
    if (greyed) await expect(entry).toBeDisabled();
    else await expect(entry).toBeEnabled();
  }
  const before = await readWorkbookState(page);
  await items.nth(1).click();
  await expect(menu).toBeHidden();
  expect(await readWorkbookState(page)).toEqual(before);
};

/** On a protected sheet the whole Table Design tab is unavailable, the table name and every style
 *  option included, and so are Insert Table and Ctrl+T. Refresh stays enabled, before and under
 *  protection, with the same menu. */
export async function runProtectedSheetTableScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform });
  await sheet.expectNoStub();
  await sheet.seed({
    A1: 'Item',
    B1: 'Qty',
    A2: 'Pen',
    B2: 3,
    A3: 'Ink',
    B3: 5,
    D5: 'Code',
    E5: 'Count',
    D6: 'X1',
    E6: 7,
  });
  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    instance.workbook.sheet(0).addTable({ range: 'A1:B3', name: 'Sales' });
    instance.history.clear();
  });
  await sheet.goTo('A2');
  await expectRefreshMenu(page);

  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).protect(),
  );
  const before = await readTables(page);
  expect(before).toHaveLength(1);

  await sheet.goTo('A2');
  await selectRibbonTab(page, 'tableDesign');
  await expect(page.locator('[data-ribbon-input="tableName"] input').first()).toBeDisabled();
  for (const command of TABLE_STYLE_OPTIONS)
    await expect(page.locator(`[data-ribbon-command="${command}"]`).first()).toBeDisabled();
  expect(await readTables(page)).toEqual(before);
  await expectRefreshMenu(page);

  if (platform === 'default') {
    await selectRibbonTab(page, 'insert');
    await expect(page.locator('[data-ribbon-command="formatTableInsert"]').first()).toBeDisabled();
  }

  // D5:E6 is a region Ctrl+T would turn into a table on an unprotected sheet.
  await sheet.goTo('D5');
  await sheet.shortcut('t');
  await page.waitForTimeout(300);
  expect(await readTables(page)).toHaveLength(1);
  await sheet.expectNoConsoleErrors();
}

const SHEET_PROTECTED_JA =
  '変更しようとしているセルやグラフは保護されているシート上にあります。変更するには、シートの保護を解除してください。パスワードの入力が必要な場合もあります。';

/** Each sheet's table and the style it uses, sheet by sheet. */
const readTableStyles = (page: Page) =>
  page.evaluate(() => {
    const workbook = (window as unknown as InstanceWindow).__fcInst.workbook;
    return [0, 1].map((index) =>
      workbook
        .sheet(index)
        .listTables()
        .map((table) => [table.name, table.getStyleOptions()?.styleName ?? null]),
    );
  });

/** The active table's Table Design strip chip for the style labelled `label`. */
const styleChip = async (page: Page, label: string) => {
  await selectRibbonTab(page, 'tableDesign');
  const chip = page.locator(`[data-table-quick-style][aria-label="${label}"]`).first();
  await revealRibbonCommand(chip);
  return chip;
};

/** Opens the active table's Quick Styles gallery. */
const openQuickStyles = async (page: Page) => {
  await selectRibbonTab(page, 'tableDesign');
  const quickStyles = page.locator('[data-ribbon-command="tableQuickStyles"]').first();
  await revealRibbonCommand(quickStyles);
  await quickStyles.click();
  const gallery = page.locator('#menu-table-quick-styles');
  await expect(gallery).toBeVisible();
  return gallery;
};

/** Right-clicks the style's chip and returns its context menu. */
const openStyleMenu = async (page: Page, label: string) => {
  await (await styleChip(page, label)).click({ button: 'right' });
  const menu = page.locator('.fc-tb__tablestyle-ctx');
  await expect(menu).toBeVisible();
  return {
    menu,
    item: (action: string) => menu.locator(`[data-table-style-context="${action}"]`),
  };
};

/** A custom table style cannot be modified or deleted while a table using it sits on a protected
 *  sheet, even from another, unprotected sheet: its strip menu greys Modify... and Delete with the
 *  protection reason, Duplicate stays available, and the style and its tables stay as they are. */
export async function runProtectedSheetCustomStyleScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Item', B1: 'Qty', A2: 'Pen', B2: 3, A3: 'Ink', B3: 5 });
  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    const workbook = instance.workbook;
    workbook.sheet(0).addTable({ range: 'A1:B3', name: 'Sales' });
    const second = workbook.addSheet('Notes');
    workbook.batch(() => {
      second.setText(0, 0, 'Code');
      second.setText(0, 1, 'Count');
      second.setText(1, 0, 'X1');
      second.setNumber(1, 1, 7);
    });
    second.addTable({ range: 'A1:B2', name: 'Codes' });
    instance.history.clear();
  });

  // A custom style made through New Table Style, then applied to both tables from the gallery.
  await sheet.goTo('A2');
  await (await openQuickStyles(page))
    .locator('[data-table-style-footer="new-table-style"]')
    .click();
  const dialog = page.locator('.fc-tablestyle-dlg');
  await expect(dialog).toBeVisible();
  await dialog.locator('[data-dialog-field="name"]').fill('Review');
  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();
  const applyReview = async () =>
    (await openQuickStyles(page)).locator('[data-table-style][aria-label="Review"]').click();
  await applyReview();
  await sheet.chooseSheet('Notes');
  await sheet.goTo('A2');
  await applyReview();
  const styled = [[['Sales', 'Review']], [['Codes', 'Review']]];
  await expect.poll(() => readTableStyles(page)).toEqual(styled);

  // Control: with no sheet protected, Modify... and Delete are available.
  let open = await openStyleMenu(page, 'Review');
  await expect(open.item('modify')).toBeEnabled();
  await expect(open.item('delete')).toBeEnabled();
  await page.keyboard.press('Escape');
  await expect(open.menu).toBeHidden();

  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    instance.workbook.sheet(0).protect();
    instance.history.clear();
  });
  open = await openStyleMenu(page, 'Review');
  for (const action of ['modify', 'delete']) {
    await expect(open.item(action), action).toBeDisabled();
    await expect(open.item(action), action).toHaveAttribute(
      'title',
      expect.stringContaining(SHEET_PROTECTED_JA),
    );
  }
  await expect(open.item('duplicate')).toBeEnabled();
  await open.item('delete').click({ force: true });
  await page.keyboard.press('Escape');
  await expect(open.menu).toBeHidden();

  expect(await readTableStyles(page)).toEqual(styled);
  await expect(await styleChip(page, 'Review')).toBeVisible();
  expect(
    await page.evaluate(() => (window as unknown as InstanceWindow).__fcInst.history.canUndo),
  ).toBe(false);
  await sheet.expectNoConsoleErrors();
}
