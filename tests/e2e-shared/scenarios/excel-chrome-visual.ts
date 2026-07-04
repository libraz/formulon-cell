import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

const snapshotDemo = async (page: Page, name: string): Promise<void> => {
  await expect(page.locator('.demo')).toHaveScreenshot(`${name}.png`, {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
  });
};

const snapshotPage = async (page: Page, name: string): Promise<void> => {
  await expect(page).toHaveScreenshot(`${name}.png`, {
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
  await page.evaluate(({ command, menuId }) => {
    const button = document.querySelector<HTMLButtonElement>(`[data-ribbon-command="${command}"]`);
    const toolbar = (
      window as unknown as {
        __fcToolbar?: {
          dropdownsApi?: {
            openDynamicRibbonDropdown?: (
              spec: { command: string; menuId: string },
              button?: HTMLButtonElement | null,
            ) => void;
          } | null;
        } | null;
      }
    ).__fcToolbar;
    if (!button || !toolbar?.dropdownsApi?.openDynamicRibbonDropdown) {
      throw new Error(`window.__fcToolbar.dropdownsApi is required to open ${command}`);
    }
    toolbar.dropdownsApi.openDynamicRibbonDropdown({ command, menuId }, button);
  }, opts);
  const menu = page.locator(`#${opts.menuId}`);
  await expect(menu).toBeVisible();
  await snapshotPage(page, opts.name);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
};

const seedTableSelection = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    const w = window as unknown as {
      __fcInst?: {
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          setNumber?: (addr: { sheet: number; row: number; col: number }, value: number) => void;
        };
        store?: {
          setState?: (fn: (state: Record<string, unknown>) => Record<string, unknown>) => void;
        };
      };
    };
    const inst = w.__fcInst;
    if (!inst?.workbook || !inst.store?.setState) {
      throw new Error('window.__fcInst with workbook/store APIs is required');
    }
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 0 }, 'Region');
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 1 }, 'Sales');
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 0 }, 'East');
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 1 }, 120);
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 0 }, 'West');
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 1 }, 80);
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...(state.selection as Record<string, unknown>),
        active: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      },
    }));
  });
};

const seedPivotFilterSelection = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    const w = window as unknown as {
      __fcInst?: {
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          setNumber?: (addr: { sheet: number; row: number; col: number }, value: number) => void;
        };
        store?: {
          setState?: (fn: (state: Record<string, unknown>) => Record<string, unknown>) => void;
        };
      };
    };
    const inst = w.__fcInst;
    if (!inst?.workbook || !inst.store?.setState) {
      throw new Error('window.__fcInst with workbook/store APIs is required');
    }
    ['Region', 'Sales', 'Qty', 'Channel', 'Segment'].forEach((value, col) => {
      inst.workbook?.setText?.({ sheet: 0, row: 0, col }, value);
    });
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 0 }, 'East');
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 1 }, 10);
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 2 }, 2);
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 3 }, 'Online');
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 4 }, 'Consumer');
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 0 }, 'West');
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 1 }, 20);
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 2 }, 4);
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 3 }, 'Retail');
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 4 }, 'Business');
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...(state.selection as Record<string, unknown>),
        active: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 },
      },
    }));
  });
};

const seedMultiSelection = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    const w = window as unknown as {
      __fcInst?: {
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          setNumber?: (addr: { sheet: number; row: number; col: number }, value: number) => void;
        };
        store?: {
          setState?: (fn: (state: Record<string, unknown>) => Record<string, unknown>) => void;
        };
      };
    };
    const inst = w.__fcInst;
    if (!inst?.workbook || !inst.store?.setState) {
      throw new Error('window.__fcInst with workbook/store APIs is required');
    }
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 0 }, 'North');
    inst.workbook.setNumber?.({ sheet: 0, row: 0, col: 1 }, 120);
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 0 }, 'South');
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 1 }, 80);
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 0 }, 'East');
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 1 }, 140);
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...(state.selection as Record<string, unknown>),
        active: { sheet: 0, row: 0, col: 0 },
        anchor: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      },
    }));
  });
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
  await expect.poll(() => formulaBar.evaluate((el) => el.clientHeight)).toBeGreaterThanOrEqual(30);
  await expect.poll(() => sheetbar.evaluate((el) => el.clientHeight)).toBeGreaterThanOrEqual(28);
  await expect.poll(() => statusbar.evaluate((el) => el.clientHeight)).toBeGreaterThanOrEqual(24);
  await expect
    .poll(() =>
      formulaBar.evaluate((el) =>
        getComputedStyle(el).getPropertyValue('--fc-formulabar-namebox-width').trim(),
      ),
    )
    .toBe('110px');
  await expect
    .poll(() =>
      sheetbar.evaluate((el) =>
        getComputedStyle(el).getPropertyValue('--fc-sheetbar-tab-height').trim(),
      ),
    )
    .toBe('28px');
};

export async function runExcelChromeVisualScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.mount({ locale: 'ja' });
  await sp.expectNoStub();
  await expectWorkbookSurfaceChrome(page);

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
  await snapshotPage(page, 'excel-chrome-home-conditional-formatting-menu-1440');
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
  await snapshotPage(page, 'excel-chrome-cell-context-paste-special-submenu-1440');
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
  await formatCells.locator('.fc-fmtdlg__footer').getByText('キャンセル', { exact: true }).click();
  await expect(formatCells).toBeHidden();

  await seedTableSelection(page);
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
  await page.locator('[data-ribbon-command="pivotTableInsert"]').click();
  const pivotFilterSource = page.getByRole('dialog', { name: 'ピボットテーブルの作成' });
  await expect(pivotFilterSource).toBeVisible();
  await pivotFilterSource.locator('[data-pivot-field-list-field="Segment"]').check();
  await pivotFilterSource.locator('[data-pivot-field-list-field="Channel"]').check();
  await pivotFilterSource.getByRole('button', { name: 'フィールドの設定: Channel' }).click();
  await expect(pivotFilterSource.locator('.fc-pivotdlg__area-settings-panel')).toContainText(
    'フィールドの設定: Channel',
  );
  await pivotFilterSource.getByRole('button', { name: 'フィルター...' }).click();
  const pivotFilter = page.getByRole('dialog', { name: 'ピボットテーブル フィルター: Channel' });
  await expect(pivotFilter).toBeVisible();
  await expect(pivotFilter.locator('[data-pivot-filter-category="true"]')).toHaveValue('label');
  await snapshotPage(page, 'excel-chrome-insert-pivottable-filter-dialog-1440');
  await pivotFilter.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(pivotFilter).toBeHidden();
  await pivotFilterSource.getByRole('button', { name: 'キャンセル' }).click();
  await expect(pivotFilterSource).toBeHidden();

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
  const nameManager = page.getByRole('dialog', { name: '名前 マネージャー' });
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
  await watchWindow.getByRole('button', { name: 'ウォッチを追加', exact: true }).click();
  await expect(watchWindow.locator('.fc-watch__row')).toHaveCount(1);
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
  await dataValidation.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(dataValidation).toBeHidden();
  await page.locator('[data-ribbon-command="textToColumns"]').click();
  const textToColumns = page.getByRole('dialog', { name: '区切り位置指定ウィザード' });
  await expect(textToColumns).toBeVisible();
  await snapshotPage(page, 'excel-chrome-data-text-to-columns-dialog-1440');
  await textToColumns.getByRole('button', { name: 'キャンセル' }).click();
  await expect(textToColumns).toBeHidden();
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
  await expect(printTitlesSetup.getByRole('tab', { name: 'シート', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await snapshotPage(page, 'excel-chrome-page-layout-print-titles-dialog-1440');
  await printTitlesSetup.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(printTitlesSetup).toBeHidden();
  await page.locator('[data-ribbon-command="selectionPanePageLayout"]').click();
  const workbookObjects = page.getByRole('dialog', { name: 'ブック オブジェクト' });
  await expect(workbookObjects).toBeVisible();
  await snapshotPage(page, 'excel-chrome-page-layout-selection-pane-1440');
  await workbookObjects.getByRole('button', { name: '閉じる', exact: true }).click();
  await expect(workbookObjects).toBeHidden();

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
  await switchRibbonTab(page, 'review');
  await snapshotDemo(page, 'excel-chrome-review-full-1440');
  await snapshotRibbonSecondaryMenu(page, {
    command: 'deleteCommentReview',
    menuId: 'menu-review-comments',
    name: 'excel-chrome-review-comments-menu-1440',
    activation: 'splitPrimary',
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
  const accessibility = page.getByRole('dialog', { name: 'アクセシビリティ チェック' });
  await expect(accessibility).toBeVisible();
  await snapshotPage(page, 'excel-chrome-review-accessibility-dialog-1440');
  await accessibility.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(accessibility).toBeHidden();
  await page.locator('[data-ribbon-command="translateReview"]').click();
  const translate = page.getByRole('dialog', { name: '翻訳' });
  await expect(translate).toBeVisible();
  await snapshotPage(page, 'excel-chrome-review-translate-dialog-1440');
  await translate.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(translate).toBeHidden();
  await page.locator('[data-ribbon-command="newCommentReview"]').click();
  const note = page.getByRole('dialog', { name: 'メモを挿入' });
  await expect(note).toBeVisible();
  await snapshotPage(page, 'excel-chrome-review-insert-note-dialog-1440');
  await note.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await expect(note).toBeHidden();
  await page.evaluate(() => {
    const inst = (
      window as unknown as {
        __fcInst?: {
          workbook?: {
            setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          };
        };
      }
    ).__fcInst;
    if (!inst?.workbook?.setText) {
      throw new Error('window.__fcInst.workbook.setText is required');
    }
    inst.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'teh teh');
  });
  await page.locator('[data-ribbon-command="spellingReview"]').click();
  const spelling = page.getByRole('dialog', { name: 'スペル チェック' });
  await expect(spelling).toBeVisible();
  await snapshotPage(page, 'excel-chrome-review-spelling-dialog-1440');
  await spelling.getByRole('button', { name: 'OK' }).click();
  await expect(spelling).toBeHidden();
  await switchRibbonTab(page, 'help');
  await snapshotDemo(page, 'excel-chrome-help-full-1440');
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

  await page.setViewportSize({ width: 390, height: 844 });
  await setRibbonDisplayMode(page, 'full');
  await expectWorkbookSurfaceChrome(page);
  await snapshotDemo(page, 'excel-chrome-home-full-mobile-390');
}
