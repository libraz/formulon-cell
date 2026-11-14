import { expect, type Page, test } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

interface SnapshotOptions {
  /** Leave the pointer where it is: the shot shows a submenu opened by hover. */
  keepPointer?: boolean;
}

/** Move the pointer off the page and wait out the ribbon screen tip, which
 *  any `[title]` control under a resting pointer opens after its delay. */
const parkPointer = async (page: Page): Promise<void> => {
  await page.mouse.move(-1, -1);
  await expect(page.locator('.fc-tb__screen-tip:visible')).toHaveCount(0);
  await page.evaluate(() => new Promise((resolve) => requestAnimationFrame(resolve)));
};

const snapshotDemo = async (
  page: Page,
  name: string,
  opts: SnapshotOptions = {},
): Promise<void> => {
  if (!opts.keepPointer) await parkPointer(page);
  // Soft, so one walk reports every shot that moved instead of stopping at the first.
  await expect.soft(page.locator('.demo')).toHaveScreenshot(`${name}.png`, {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
  });
};

const snapshotPage = async (
  page: Page,
  name: string,
  opts: SnapshotOptions = {},
): Promise<void> => {
  if (!opts.keepPointer) await parkPointer(page);
  await expect.soft(page).toHaveScreenshot(`${name}.png`, {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
  });
};

const switchRibbonTab = async (page: Page, tab: string): Promise<void> => {
  await page.locator(`[data-ribbon-tab="${tab}"]`).click();
  await expect(page.locator(`[data-ribbon-tab="${tab}"][aria-selected="true"]`)).toHaveCount(1);
};

const snapshotRibbonMenu = async (
  page: Page,
  opts: {
    command: string;
    menuId: string;
    name: string;
    activation?: 'dropdown' | 'gallery';
  },
): Promise<void> => {
  const activation = opts.activation ?? 'dropdown';
  const button = page.locator(`[data-ribbon-command="${opts.command}"]`).first();
  await expect(button).toHaveAttribute('data-ribbon-menu-id', opts.menuId);
  await expect(button).toHaveAttribute('data-ribbon-activation', activation);
  await button.click();
  const menu = page.locator(`#${opts.menuId}`);
  await expect(menu).toBeVisible();
  await snapshotPage(page, opts.name);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

const snapshotRibbonSecondaryMenu = async (
  page: Page,
  opts: {
    command: string;
    menuId: string;
    name: string;
    activation: 'splitPrimary' | 'splitToggle';
  },
): Promise<void> => {
  const button = page.locator(`[data-ribbon-command="${opts.command}"]`).first();
  await expect(button).toHaveAttribute('data-ribbon-menu-id', opts.menuId);
  await expect(button).toHaveAttribute('data-ribbon-activation', opts.activation);
  // The primary face runs the command; the split chevron opens the secondary menu.
  await button.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator(`#${opts.menuId}`);
  await expect(menu).toBeVisible();
  await snapshotPage(page, opts.name);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

const seedTableSelection = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      sheet.setText(0, 0, 'Region');
      sheet.setText(0, 1, 'Sales');
      sheet.setText(1, 0, 'East');
      sheet.setNumber(1, 1, 120);
      sheet.setText(2, 0, 'West');
      sheet.setNumber(2, 1, 80);
    });
    inst.history.clear();
    inst.selection.set('A1:B3');
  });
};

const seedPivotFilterSelection = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      ['Region', 'Sales', 'Qty', 'Channel', 'Segment'].forEach((value, col) => {
        sheet.setText(0, col, value);
      });
      sheet.setText(1, 0, 'East');
      sheet.setNumber(1, 1, 10);
      sheet.setNumber(1, 2, 2);
      sheet.setText(1, 3, 'Online');
      sheet.setText(1, 4, 'Consumer');
      sheet.setText(2, 0, 'West');
      sheet.setNumber(2, 1, 20);
      sheet.setNumber(2, 2, 4);
      sheet.setText(2, 3, 'Retail');
      sheet.setText(2, 4, 'Business');
    });
    inst.history.clear();
    inst.selection.set('A1:E3');
  });
};

/** Builds a report at G1 over the seeded rows, Channel on rows and Sum of Sales, as the Create
 *  PivotTable dialog does: place it, refresh the sheet's cells and fit the report's columns. */
const createChannelReport = (page: Page): Promise<void> =>
  page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const workbook = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    const made = w.__fcProbes['commands/pivot-table'].createPivotTableAtPlacement(
      { store, workbook, history: null, origin: 'instanceApi', commandId: 'pivotTableInsert' },
      {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 },
        name: 'ChannelReport',
        rowField: 'Channel',
        valueField: 'Sales',
        aggregation: 0,
      },
      { row: 0, col: 6 },
    );
    if (!made.ok) throw new Error(`pivot creation failed: ${made.reason}`);
    w.__fcProbes['store/store'].mutators.refreshActiveSheetCells(store, workbook);
    w.__fcProbes['commands/pivot-table'].autofitPivotTableColumns(store, workbook);
    w.__fcInst.history.clear();
  });

const seedMultiSelection = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      sheet.setText(0, 0, 'North');
      sheet.setNumber(0, 1, 120);
      sheet.setText(1, 0, 'South');
      sheet.setNumber(1, 1, 80);
      sheet.setText(2, 0, 'East');
      sheet.setNumber(2, 1, 140);
    });
    inst.history.clear();
    inst.selection.set('A1:B3');
  });
};

/** Collapses the selection onto a single cell. */
const selectSingleCell = async (page: Page, cell: { row: number; col: number }): Promise<void> => {
  await page.evaluate(({ row, col }) => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.selection.set(`${String.fromCharCode(65 + col)}${row + 1}`);
  }, cell);
};

const setRibbonDisplayMode = async (page: Page, mode: string): Promise<void> => {
  await page.locator('[data-ribbon-toggle]').click();
  await page.locator(`[data-ribbon-display-option="${mode}"]`).click();
  await expect(page.locator(`.fc-tb__ribbon-shell--${mode}`)).toHaveCount(1);
};

const snapshotRibbonDisplayOptionsMenu = async (page: Page): Promise<void> => {
  const toggle = page.locator('[data-ribbon-toggle]').first();
  await expect(toggle).toHaveAttribute('aria-label', 'リボンの表示オプション');
  await toggle.click();
  const menu = page.locator('.fc-tb__ribbon-display-menu');
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitemradio', { name: 'リボンを常に表示' })).toHaveAttribute(
    'aria-checked',
    'true',
  );
  await expect(menu.getByRole('menuitemradio', { name: '1 行のリボン' })).toBeVisible();
  await expect(menu.getByRole('menuitemradio', { name: 'タブのみ表示' })).toBeVisible();
  await expect(menu.getByRole('menuitemradio', { name: 'リボンを自動的に非表示' })).toBeVisible();
  await snapshotPage(page, 'excel-chrome-ribbon-display-options-menu-1024');
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

const expectWorkbookSurfaceChrome = async (page: Page): Promise<void> => {
  const formulaBar = page.locator('.fc-host__formulabar').first();
  const nameBox = page.locator('.fc-host__formulabar-tag').first();
  const formulaInput = page.locator('.fc-host__formulabar-input').first();
  const sheetbar = page.locator('.fc-host__sheetbar').first();
  const statusbar = page.locator('.fc-host__statusbar').first();
  await expect(formulaBar).toBeVisible();
  await expect(nameBox).toBeVisible();
  await expect(formulaInput).toBeVisible();
  await expect(sheetbar).toBeVisible();
  await expect(statusbar).toBeVisible();
  // Heights track the `--fc-*-height` tokens the surface stylesheets set; the
  // token assertions below pin the widths those bars are laid out against.
  // clientHeight excludes the hairline rule each bar carries, so the sheet bar
  // lands one pixel under its 28px token.
  await expect.poll(() => formulaBar.evaluate((el) => el.clientHeight)).toBeGreaterThanOrEqual(28);
  await expect.poll(() => sheetbar.evaluate((el) => el.clientHeight)).toBeGreaterThanOrEqual(27);
  await expect.poll(() => statusbar.evaluate((el) => el.clientHeight)).toBeGreaterThanOrEqual(24);
  await expect
    .poll(() =>
      formulaBar.evaluate((el) =>
        getComputedStyle(el).getPropertyValue('--fc-formulabar-namebox-width').trim(),
      ),
    )
    .toBe('86px');
  await expect
    .poll(() =>
      sheetbar.evaluate((el) =>
        getComputedStyle(el).getPropertyValue('--fc-sheetbar-tab-height').trim(),
      ),
    )
    .toBe('26px');
};

export async function runExcelChromeVisualScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  // A control the walk no longer finds fails its step within seconds, not at the test timeout.
  page.setDefaultTimeout(15_000);
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.mount({ locale: 'ja' });
  await sp.expectNoStub();
  await expectWorkbookSurfaceChrome(page);

  await test.step('home tab', async () => {
    await switchRibbonTab(page, 'home');
    await snapshotDemo(page, 'excel-chrome-home-full-1440');
    await seedMultiSelection(page);
    await snapshotDemo(page, 'excel-chrome-home-multi-selection-1440');
    await sp.focusHost();
    await page.keyboard.type('Editing');
    await expect(page.locator('.fc-host__editor')).toBeVisible();
    await snapshotDemo(page, 'excel-chrome-home-cell-editing-1440');
    await page.keyboard.press('Escape');
    await expect(page.locator('.fc-host__editor')).toBeHidden();
    await page.locator('.fc-host__statusbar').click({ button: 'right' });
    await expect(page.locator('.fc-statusbar__chooser')).toBeVisible();
    await snapshotPage(page, 'excel-chrome-statusbar-chooser-1440');
    await page.keyboard.press('Escape');
    await expect(page.locator('.fc-statusbar__chooser')).toBeHidden();

    const search = page.getByRole('combobox', { name: 'コマンドの検索' });
    await search.fill('ヘルプ');
    const searchResults = page.locator('#demo-search-results');
    await expect(searchResults).toBeVisible();
    const disabledResult = searchResults
      .locator('.fc-tb__command-item[aria-disabled="true"]')
      .first();
    await expect(disabledResult).toBeVisible();
    await expect(disabledResult).toHaveAttribute('aria-disabled', 'true');
    await snapshotPage(page, 'excel-chrome-search-disabled-result-1440');
    await page.keyboard.press('Escape');
    await expect(searchResults).toBeHidden();

    await snapshotRibbonMenu(page, {
      command: 'formatTableHome',
      menuId: 'menu-table-style-home',
      name: 'excel-chrome-home-format-as-table-gallery-1440',
      activation: 'gallery',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'paste',
      menuId: 'menu-paste',
      name: 'excel-chrome-home-paste-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'underline',
      menuId: 'menu-underline',
      name: 'excel-chrome-home-underline-menu-1440',
      activation: 'splitToggle',
    });
    await page.locator('[data-ribbon-command="conditional"]').click();
    await expect(page.locator('#menu-conditional')).toBeVisible();
    await page.locator('#menu-conditional [data-cf-submenu="dataBar"]').hover();
    await expect(page.locator('#menu-conditional [data-cf-panel="dataBar"]')).toBeVisible();
    await snapshotPage(page, 'excel-chrome-home-conditional-formatting-menu-1440', {
      keepPointer: true,
    });
    await page.keyboard.press('Escape');
    await page.locator('[data-ribbon-command="conditional"]').click();
    await expect(page.locator('#menu-conditional')).toBeVisible();
    await page.locator('#menu-conditional [data-cf-submenu="highlight"]').hover();
    await expect(page.locator('#menu-conditional-highlight')).toBeVisible();
    await page.locator('#menu-conditional-highlight [data-cf-action="cell-gt"]').click();
    const greaterThanDialog = page.getByRole('dialog', { name: '指定の値より大きい...' });
    await expect(greaterThanDialog).toBeVisible();
    await snapshotPage(page, 'excel-chrome-home-conditional-greater-than-dialog-1440');
    await greaterThanDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(greaterThanDialog).toBeHidden();
    await snapshotRibbonMenu(page, {
      command: 'textOrientation',
      menuId: 'menu-text-orientation',
      name: 'excel-chrome-home-text-orientation-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'borders',
      menuId: 'menu-borders',
      name: 'excel-chrome-home-borders-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'fillHome',
      menuId: 'menu-fill',
      name: 'excel-chrome-home-fill-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'cellStyles',
      menuId: 'menu-cell-styles-home',
      name: 'excel-chrome-home-cell-styles-gallery-1440',
      activation: 'gallery',
    });
    await snapshotRibbonMenu(page, {
      command: 'insertRows',
      menuId: 'menu-insert-cells',
      name: 'excel-chrome-home-insert-cells-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'deleteRows',
      menuId: 'menu-delete-cells',
      name: 'excel-chrome-home-delete-cells-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'formatCellsHome',
      menuId: 'menu-format-cells',
      name: 'excel-chrome-home-format-cells-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'clearFormat',
      menuId: 'menu-clear',
      name: 'excel-chrome-home-clear-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'sortFilterHome',
      menuId: 'menu-sort-home',
      name: 'excel-chrome-home-sort-filter-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'findHome',
      menuId: 'menu-find-select',
      name: 'excel-chrome-home-find-select-menu-1440',
    });
    await page.locator('[data-ribbon-command="findHome"]').click();
    await expect(page.locator('#menu-find-select')).toBeVisible();
    await page.locator('#menu-find-select [data-find-select="find"]').click();
    const findReplace = page.locator('.fc-find');
    await expect(findReplace).toBeVisible();
    await snapshotPage(page, 'excel-chrome-home-find-replace-dialog-1440');
    await findReplace.getByRole('button', { name: 'オプション >>', exact: true }).click();
    await expect(
      findReplace.getByRole('button', { name: 'オプション <<', exact: true }),
    ).toBeVisible();
    await snapshotPage(page, 'excel-chrome-home-find-replace-options-1440');
    await findReplace.locator('.fc-find__btn--close').click();
    await expect(findReplace).toBeHidden();
    await page.locator('[data-ribbon-command="findHome"]').click();
    await expect(page.locator('#menu-find-select')).toBeVisible();
    await page.locator('#menu-find-select [data-find-select="replace"]').click();
    await expect(findReplace).toBeVisible();
    await expect(findReplace.locator('.fc-find__tab[aria-selected="true"]')).toHaveText('置換');
    await snapshotPage(page, 'excel-chrome-home-replace-dialog-1440');
    await findReplace.locator('.fc-find__btn--close').click();
    await expect(findReplace).toBeHidden();
    await page.locator('[data-ribbon-command="findHome"]').click();
    await expect(page.locator('#menu-find-select')).toBeVisible();
    await page.locator('#menu-find-select [data-find-select="go-to-special"]').click();
    const goToSpecial = page.getByRole('dialog', { name: '選択オプション' });
    await expect(goToSpecial).toBeVisible();
    await snapshotPage(page, 'excel-chrome-home-go-to-special-dialog-1440');
    await goToSpecial.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(goToSpecial).toBeHidden();

    await sp.focusHost();
    await sp.shortcut('c');
    await page
      .locator('.fc-host')
      .first()
      .click({ button: 'right', position: { x: 200, y: 200 } });
    const contextMenu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
    await expect(contextMenu).toBeVisible();
    await snapshotPage(page, 'excel-chrome-cell-context-menu-1440');
    await contextMenu.locator('[data-fc-submenu="pasteSpecialMenu"]').hover();
    const pasteSpecialSubmenu = page.locator('.fc-ctxmenu__sub');
    await expect(pasteSpecialSubmenu).toBeVisible();
    await snapshotPage(page, 'excel-chrome-cell-context-paste-special-submenu-1440', {
      keepPointer: true,
    });
    await pasteSpecialSubmenu.locator('[data-fc-action="pasteSpecial"]').click();
    await expect(contextMenu).toBeHidden();
    const pasteSpecialDialog = page.getByRole('dialog', { name: '形式を選択して貼り付け' });
    await expect(pasteSpecialDialog).toBeVisible();
    await snapshotPage(page, 'excel-chrome-paste-special-dialog-1440');
    await pasteSpecialDialog
      .locator('.fc-fmtdlg__footer')
      .getByText('キャンセル', { exact: true })
      .click();
    await expect(pasteSpecialDialog).toBeHidden();
    await page
      .locator('.fc-host')
      .first()
      .click({ button: 'right', position: { x: 200, y: 200 } });
    await expect(contextMenu).toBeVisible();
    await contextMenu.getByText('セルの書式設定…', { exact: true }).first().click();
    await expect(contextMenu).toBeHidden();
    const formatCells = page.getByRole('dialog', { name: 'セルの書式設定' });
    await expect(formatCells).toBeVisible();
    await snapshotPage(page, 'excel-chrome-format-cells-dialog-1440');
    await formatCells
      .locator('.fc-fmtdlg__footer')
      .getByText('キャンセル', { exact: true })
      .click();
    await expect(formatCells).toBeHidden();

    await seedTableSelection(page);
  });

  await test.step('insert tab', async () => {
    await switchRibbonTab(page, 'insert');
    await snapshotDemo(page, 'excel-chrome-insert-full-1440');
    await snapshotRibbonSecondaryMenu(page, {
      command: 'pivotTableInsert',
      menuId: 'menu-pivot-table',
      name: 'excel-chrome-insert-pivottable-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonMenu(page, {
      command: 'pictureInsert',
      menuId: 'menu-picture-insert',
      name: 'excel-chrome-insert-pictures-menu-1440',
      activation: 'gallery',
    });
    await snapshotRibbonMenu(page, {
      command: 'shapesInsert',
      menuId: 'menu-shapes-insert',
      name: 'excel-chrome-insert-shapes-menu-1440',
      activation: 'gallery',
    });
    await snapshotRibbonMenu(page, {
      command: 'screenshotInsert',
      menuId: 'menu-screenshot-insert',
      name: 'excel-chrome-insert-screenshot-menu-1440',
      activation: 'gallery',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'chartInsert',
      menuId: 'menu-chart-insert',
      name: 'excel-chrome-insert-chart-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'symbolInsert',
      menuId: 'menu-symbol',
      name: 'excel-chrome-insert-symbol-menu-1440',
      activation: 'splitPrimary',
    });
    await page.locator('[data-ribbon-command="formatTableInsert"]').click();
    const createTable = page.getByRole('dialog', { name: 'テーブルの作成' });
    await expect(createTable).toBeVisible();
    await snapshotPage(page, 'excel-chrome-insert-create-table-dialog-1440');
    await createTable.getByRole('button', { name: 'キャンセル' }).click();
    await expect(createTable).toBeHidden();
    await page.locator('[data-ribbon-command="pivotTableInsert"]').click();
    const createPivotTable = page.getByRole('dialog', { name: 'ピボットテーブルの作成' });
    await expect(createPivotTable).toBeVisible();
    await snapshotPage(page, 'excel-chrome-insert-create-pivottable-dialog-1440');
    await createPivotTable.getByRole('button', { name: 'キャンセル' }).click();
    await expect(createPivotTable).toBeHidden();
    await seedPivotFilterSelection(page);
    await createChannelReport(page);
    // A row label's Filter > Label Filters... opens the panel; a label condition opens the dialog.
    const label = await gridCellCenter(page, 1, 6);
    await page.mouse.click(label.x, label.y, { button: 'right' });
    await page.locator('.fc-ctxmenu [data-fc-action="filter"]').first().hover();
    await page.locator('.fc-ctxmenu__sub [data-fc-action="pivotFilterLabel"]').click();
    const pivotFilterPanel = page.locator('.fc-pivot-filter-panel');
    await expect(pivotFilterPanel).toBeVisible();
    await snapshotPage(page, 'excel-chrome-pivot-filter-panel-1440');
    // 12 is PivotFilterKind.LabelEqual.
    await pivotFilterPanel.locator('select[data-pivot-filter-by="label"]').selectOption('12');
    const pivotFilter = page.getByRole('dialog', { name: 'ラベル フィルター (Channel)' });
    await expect(pivotFilter).toBeVisible();
    // The panel closes as it hands over to the dialog.
    await expect(pivotFilterPanel).toBeHidden();
    await snapshotPage(page, 'excel-chrome-pivot-label-filter-dialog-1440');
    await pivotFilter.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(pivotFilter).toBeHidden();
    // The report's field list holds the side pane later steps open into.
    const fieldList = page.getByRole('dialog', { name: 'ピボットテーブルのフィールド' });
    await fieldList.getByRole('button', { name: '閉じる', exact: true }).click();
    await expect(fieldList).toBeHidden();
    // Later shots show the sheet without the report.
    await page.evaluate(() => {
      const inst = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      inst.workbook.sheet(0).removePivot('ChannelReport');
      inst.history.clear();
      inst.selection.set('A1');
    });
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (
            window as unknown as { __fcInst: SpreadsheetInstance }
          ).__fcInst.workbook.sheet(0);
          return [sheet.getValue(0, 6), sheet.getValue(1, 6), sheet.getValue(0, 7)].map((value) =>
            value === undefined || value === '' ? null : value,
          );
        }),
      )
      .toEqual([null, null, null]);
  });

  await test.step('formulas tab', async () => {
    await switchRibbonTab(page, 'formulas');
    await snapshotDemo(page, 'excel-chrome-formulas-full-1440');
    await snapshotRibbonSecondaryMenu(page, {
      command: 'autosumFormula',
      menuId: 'menu-autosum-formulas',
      name: 'excel-chrome-formulas-autosum-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'namedRanges',
      menuId: 'menu-defined-names',
      name: 'excel-chrome-formulas-defined-names-menu-1440',
      activation: 'splitPrimary',
    });
    await page.locator('[data-ribbon-command="namedRanges"]').click();
    const nameManager = page.getByRole('dialog', { name: '名前の管理' });
    await expect(nameManager).toBeVisible();
    await snapshotPage(page, 'excel-chrome-formulas-name-manager-dialog-1440');
    await nameManager.getByRole('button', { name: '閉じる', exact: true }).click();
    await expect(nameManager).toBeHidden();
    await snapshotRibbonMenu(page, {
      command: 'clearArrows',
      menuId: 'menu-clear-arrows',
      name: 'excel-chrome-formulas-clear-arrows-menu-1440',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'errorChecking',
      menuId: 'menu-error-checking',
      name: 'excel-chrome-formulas-error-checking-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'watch',
      menuId: 'menu-watch-formulas',
      name: 'excel-chrome-formulas-watch-menu-1440',
      activation: 'splitPrimary',
    });
    await page.locator('[data-ribbon-command="watch"]').click();
    const watchWindow = page.locator('.fc-watch');
    await expect(watchWindow).toBeVisible();
    // Add Watch takes the whole selection, so collapse it onto one untouched
    // cell first — otherwise the row count tracks whatever is selected here.
    await selectSingleCell(page, { row: 5, col: 3 });
    const watchRows = watchWindow.locator('.fc-watch__row');
    const rowsBeforeAdd = await watchRows.count();
    await watchWindow.getByRole('button', { name: 'ウォッチを追加', exact: true }).click();
    await expect(watchRows).toHaveCount(rowsBeforeAdd + 1);
    await snapshotPage(page, 'excel-chrome-formulas-watch-window-1440');
    await watchWindow.getByRole('button', { name: '閉じる', exact: true }).click();
    await expect(watchWindow).toBeHidden();
    await snapshotRibbonMenu(page, {
      command: 'calcOptions',
      menuId: 'menu-calc-options',
      name: 'excel-chrome-formulas-calc-options-menu-1440',
    });
    await page.locator('[data-ribbon-command="fx"]').click();
    const functionArguments = page.getByRole('dialog', { name: '関数の引数' });
    await expect(functionArguments).toBeVisible();
    await snapshotPage(page, 'excel-chrome-formulas-function-arguments-dialog-1440');
    await functionArguments.getByRole('button', { name: 'キャンセル' }).click();
    await expect(functionArguments).toBeHidden();
  });

  await test.step('data tab', async () => {
    await switchRibbonTab(page, 'data');
    await snapshotDemo(page, 'excel-chrome-data-full-1440');
    await snapshotRibbonMenu(page, {
      command: 'filter',
      menuId: 'menu-sort',
      name: 'excel-chrome-data-sort-filter-menu-1440',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'textToColumns',
      menuId: 'menu-text-to-columns',
      name: 'excel-chrome-data-text-to-columns-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'dataValidation',
      menuId: 'menu-data-validation',
      name: 'excel-chrome-data-validation-menu-1440',
      activation: 'splitPrimary',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'linksData',
      menuId: 'menu-links-data',
      name: 'excel-chrome-data-links-menu-1440',
      activation: 'splitPrimary',
    });
    await page.locator('[data-ribbon-command="sortData"]').click();
    const sortDialog = page.getByRole('dialog', { name: '並べ替え' });
    await expect(sortDialog).toBeVisible();
    await snapshotPage(page, 'excel-chrome-data-custom-sort-dialog-1440');
    await sortDialog.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(sortDialog).toBeHidden();
    await page.locator('[data-ribbon-command="removeDupes"]').click();
    const removeDuplicates = page.getByRole('dialog', { name: '重複の削除' });
    await expect(removeDuplicates).toBeVisible();
    await snapshotPage(page, 'excel-chrome-data-remove-duplicates-dialog-1440');
    await removeDuplicates.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(removeDuplicates).toBeHidden();
    await page.locator('[data-ribbon-command="dataValidation"]').click();
    const dataValidation = page.getByRole('dialog', { name: '入力規則' });
    await expect(dataValidation).toBeVisible();
    await snapshotPage(page, 'excel-chrome-data-validation-dialog-1440');
    // The dialog's close affordance carries the same label as the footer button,
    // so target the footer one explicitly.
    await dataValidation.locator('button.fc-fmtdlg__btn', { hasText: 'キャンセル' }).click();
    await expect(dataValidation).toBeHidden();
    await page.locator('[data-ribbon-command="textToColumns"]').click();
    const textToColumns = page.getByRole('dialog', { name: '区切り位置指定ウィザード' });
    await expect(textToColumns).toBeVisible();
    await snapshotPage(page, 'excel-chrome-data-text-to-columns-dialog-1440');
    await textToColumns.getByRole('button', { name: 'キャンセル' }).click();
    await expect(textToColumns).toBeHidden();
  });

  await test.step('page layout tab', async () => {
    await switchRibbonTab(page, 'pageLayout');
    await snapshotDemo(page, 'excel-chrome-page-layout-full-1440');
    await snapshotRibbonMenu(page, {
      command: 'pageTheme',
      menuId: 'menu-page-theme',
      name: 'excel-chrome-page-layout-theme-menu-1440',
      activation: 'gallery',
    });
    await snapshotRibbonMenu(page, {
      command: 'printArea',
      menuId: 'menu-print-area',
      name: 'excel-chrome-page-layout-print-area-menu-1440',
    });
    await snapshotRibbonMenu(page, {
      command: 'pageBreaks',
      menuId: 'menu-page-breaks',
      name: 'excel-chrome-page-layout-breaks-menu-1440',
    });
    await page.locator('[data-ribbon-command="pageSetupAdvanced"]').click();
    const pageSetup = page.getByRole('dialog', { name: 'ページ設定' });
    await expect(pageSetup).toBeVisible();
    await snapshotPage(page, 'excel-chrome-page-setup-dialog-1440');
    await pageSetup.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(pageSetup).toBeHidden();
    await page.locator('[data-ribbon-command="printTitles"]').click();
    const printTitlesSetup = page.getByRole('dialog', { name: 'ページ設定' });
    await expect(printTitlesSetup).toBeVisible();
    await expect(
      printTitlesSetup.getByRole('tab', { name: 'シート', exact: true }),
    ).toHaveAttribute('aria-selected', 'true');
    await snapshotPage(page, 'excel-chrome-page-layout-print-titles-dialog-1440');
    await printTitlesSetup.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(printTitlesSetup).toBeHidden();
    await page.locator('[data-ribbon-command="selectionPanePageLayout"]').click();
    const workbookObjects = page.getByRole('dialog', { name: 'ブック オブジェクト' });
    await expect(workbookObjects).toBeVisible();
    await snapshotPage(page, 'excel-chrome-page-layout-selection-pane-1440');
    await workbookObjects.getByRole('button', { name: '閉じる', exact: true }).click();
    await expect(workbookObjects).toBeHidden();
  });

  await test.step('view tab', async () => {
    await switchRibbonTab(page, 'view');
    await snapshotDemo(page, 'excel-chrome-view-full-1440');
    await snapshotRibbonMenu(page, {
      command: 'freeze',
      menuId: 'menu-freeze',
      name: 'excel-chrome-view-freeze-menu-1440',
    });
    await page.locator('[data-ribbon-command="zoomDialog"]').click();
    const zoomDialog = page.getByRole('dialog', { name: 'ズーム' });
    await expect(zoomDialog).toBeVisible();
    await snapshotPage(page, 'excel-chrome-view-zoom-dialog-1440');
    await zoomDialog.getByRole('button', { name: 'キャンセル' }).click();
    await expect(zoomDialog).toBeHidden();
    await page.locator('[data-ribbon-command="viewPageLayout"]').click();
    await expect(page.locator('[data-ribbon-command="viewPageLayout"]')).toHaveAttribute(
      'aria-pressed',
      'true',
    );
    await snapshotDemo(page, 'excel-chrome-view-page-layout-mode-1440');
    await page.locator('[data-ribbon-command="viewPageBreakPreview"]').click();
    await expect(page.locator('[data-ribbon-command="viewPageBreakPreview"]')).toHaveAttribute(
      'aria-pressed',
      'true',
    );
    await snapshotDemo(page, 'excel-chrome-view-page-break-preview-mode-1440');
    await page.locator('[data-ribbon-command="viewR1C1"]').click();
    await expect(page.locator('[data-ribbon-command="viewR1C1"]')).toHaveAttribute(
      'aria-pressed',
      'true',
    );
    await snapshotDemo(page, 'excel-chrome-view-r1c1-reference-style-1440');
    await page.locator('[data-ribbon-command="viewR1C1"]').click();
    await page.locator('[data-ribbon-command="viewNormal"]').click();
    await expect(page.locator('[data-ribbon-command="viewNormal"]')).toHaveAttribute(
      'aria-pressed',
      'true',
    );
  });

  await test.step('review tab', async () => {
    await switchRibbonTab(page, 'review');
    await snapshotDemo(page, 'excel-chrome-review-full-1440');
    await snapshotRibbonMenu(page, {
      command: 'notesReview',
      menuId: 'menu-review-notes',
      name: 'excel-chrome-review-notes-menu-1440',
    });
    await snapshotRibbonSecondaryMenu(page, {
      command: 'protectReview',
      menuId: 'menu-protect-review',
      name: 'excel-chrome-review-protect-menu-1440',
      activation: 'splitPrimary',
    });
    await page.locator('[data-ribbon-command="protectReview"]').click();
    const protectSheet = page.getByRole('dialog', { name: 'シートを保護' });
    await expect(protectSheet).toBeVisible();
    await snapshotPage(page, 'excel-chrome-review-protect-sheet-dialog-1440');
    await protectSheet.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(protectSheet).toBeHidden();
    await page.locator('[data-ribbon-command="protectionReview"]').click();
    const allowEditRanges = page.getByRole('dialog', { name: '範囲の編集を許可' });
    await expect(allowEditRanges).toBeVisible();
    await snapshotPage(page, 'excel-chrome-review-allow-edit-ranges-dialog-1440');
    await allowEditRanges.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(allowEditRanges).toBeHidden();
    await page.locator('[data-ribbon-command="accessibility"]').click();
    const accessibility = page.getByRole('dialog', { name: 'アクセシビリティ', exact: true });
    await expect(accessibility).toBeVisible();
    await snapshotPage(page, 'excel-chrome-review-accessibility-pane-1440');
    await accessibility.getByRole('button', { name: '閉じる', exact: true }).click();
    await expect(accessibility).toHaveCount(0);
    await page.locator('[data-ribbon-command="translateReview"]').click();
    const translate = page.getByRole('dialog', { name: '翻訳' });
    await expect(translate).toBeVisible();
    await snapshotPage(page, 'excel-chrome-review-translate-dialog-1440');
    await translate.getByRole('button', { name: 'OK', exact: true }).click();
    await expect(translate).toBeHidden();
    await page.locator('[data-ribbon-command="notesReview"]').click();
    await page.locator('#menu-review-notes [data-comment-action="new-note"]').click();
    const note = page.getByRole('dialog', { name: 'メモを挿入' });
    await expect(note).toBeVisible();
    await snapshotPage(page, 'excel-chrome-review-insert-note-dialog-1440');
    await note.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(note).toBeHidden();
    await page.evaluate(() => {
      const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
      if (!inst) throw new Error('window.__fcInst is required');
      inst.workbook.batch(() => inst.workbook.sheet(0).setText(0, 0, 'teh teh'));
      inst.history.clear();
    });
    await page.locator('[data-ribbon-command="spellingReview"]').click();
    const spelling = page.getByRole('dialog', { name: /^スペル チェック/ });
    await expect(spelling).toBeVisible();
    await snapshotPage(page, 'excel-chrome-review-spelling-dialog-1440');
    await spelling.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(spelling).toBeHidden();
  });

  await test.step('help tab', async () => {
    await switchRibbonTab(page, 'help');
    await snapshotDemo(page, 'excel-chrome-help-full-1440');
  });

  await test.step('backstage', async () => {
    await switchRibbonTab(page, 'file');
    await expect(page.locator('.fc-tb__backstage[role="dialog"]')).toBeVisible();
    await snapshotDemo(page, 'excel-chrome-backstage-info-1440');
    await page
      .locator('.fc-tb__backstage')
      .getByRole('button', { name: '印刷', exact: true })
      .click();
    await snapshotDemo(page, 'excel-chrome-backstage-print-1440');
    await page.locator('.fc-tb__backstage').getByRole('button', { name: 'ページ設定' }).click();
    const backstagePageSetup = page.getByRole('dialog', { name: 'ページ設定' });
    await expect(backstagePageSetup).toBeVisible();
    await backstagePageSetup.getByRole('tab', { name: '余白', exact: true }).click();
    await expect(backstagePageSetup.getByText('ページ中央', { exact: true }).first()).toBeVisible();
    await snapshotPage(page, 'excel-chrome-backstage-page-setup-margins-dialog-1440');
    await backstagePageSetup.getByRole('button', { name: 'キャンセル', exact: true }).click();
    await expect(backstagePageSetup).toBeHidden();
    await page
      .locator('.fc-tb__backstage')
      .getByRole('button', { name: '閉じる', exact: true })
      .click();
    await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);
    await page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').click({ button: 'right' });
    await expect(page.locator('.fc-sheetmenu')).toBeVisible();
    await snapshotPage(page, 'excel-chrome-sheet-tab-context-menu-1440');
    await page.keyboard.press('Escape');
    await expect(page.locator('.fc-sheetmenu')).toBeHidden();
  });

  await test.step('ribbon display modes at 1024', async () => {
    await page.setViewportSize({ width: 1024, height: 768 });
    await switchRibbonTab(page, 'home');
    await snapshotRibbonDisplayOptionsMenu(page);
    await setRibbonDisplayMode(page, 'singleLine');
    await snapshotDemo(page, 'excel-chrome-home-single-line-1024');
    await setRibbonDisplayMode(page, 'tabsOnly');
    await snapshotDemo(page, 'excel-chrome-home-tabs-only-1024');
    await setRibbonDisplayMode(page, 'autoHide');
    await page.keyboard.press('Alt');
    await expect(page.locator('.fc-tb__ribbon-shell--autoHidePeek')).toHaveCount(1);
    await snapshotDemo(page, 'excel-chrome-home-auto-hide-peek-1024');
  });

  await test.step('mobile at 390', async () => {
    await page.setViewportSize({ width: 390, height: 844 });
    await setRibbonDisplayMode(page, 'full');
    await expectWorkbookSurfaceChrome(page);
    await snapshotDemo(page, 'excel-chrome-home-full-mobile-390');
  });

  await test.step('mac home and preferences', async () => {
    await page.setViewportSize({ width: 1440, height: 900 });
    await sp.mount({ locale: 'ja', platform: 'mac' });
    await sp.expectNoStub();
    // The book opens on Home; clicking the selected tab would collapse the ribbon.
    await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);
    await expect(page.locator('[data-ribbon-command="wrap"]').first()).toBeVisible();
    await snapshotDemo(page, 'excel-chrome-mac-home-full-1440');
    await page.locator('.demo__brand-mark').first().click();
    const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
    await expect(backstage).toBeVisible();
    await backstage.getByRole('button', { name: 'オプション', exact: true }).first().click();
    await expect(page.locator('.fc-prefs[role="dialog"]')).toBeVisible();
    await snapshotPage(page, 'excel-chrome-mac-preferences-1440');
  });
}
