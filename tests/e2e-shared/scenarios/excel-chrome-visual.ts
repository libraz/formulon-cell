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

const setRibbonDisplayMode = async (page: Page, mode: string): Promise<void> => {
  await page.locator('[data-ribbon-toggle]').click();
  await page.locator(`[data-ribbon-display-option="${mode}"]`).click();
  await expect(page.locator(`.demo__ribbon-shell--${mode}`)).toHaveCount(1);
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
  await sp.mount();
  await sp.expectNoStub();
  await expectWorkbookSurfaceChrome(page);

  await switchRibbonTab(page, 'home');
  await snapshotDemo(page, 'excel-chrome-home-full-1440');
  await page.locator('.fc-host__statusbar').click({ button: 'right' });
  await expect(page.locator('.fc-statusbar__chooser')).toBeVisible();
  await snapshotPage(page, 'excel-chrome-statusbar-chooser-1440');
  await page.keyboard.press('Escape');
  await expect(page.locator('.fc-statusbar__chooser')).toBeHidden();

  const search = page.getByRole('combobox', { name: 'Search commands' });
  await search.fill('coming soon');
  await expect(page.locator('.demo__command-list')).toBeVisible();
  const disabledResult = page.locator('.demo__command-item').filter({ hasText: 'Coming soon' });
  await expect(disabledResult).toBeVisible();
  await expect(disabledResult).toHaveAttribute('aria-disabled', 'true');
  await snapshotPage(page, 'excel-chrome-search-disabled-result-1440');
  await page.keyboard.press('Escape');
  await expect(page.locator('.demo__command-list')).toBeHidden();

  await page.locator('[data-ribbon-command="formatTableHome"]').click();
  await expect(page.locator('#menu-table-style-home')).toBeVisible();
  await snapshotPage(page, 'excel-chrome-home-format-as-table-gallery-1440');
  await page.keyboard.press('Escape');
  await page.locator('[data-ribbon-command="conditional"]').click();
  await expect(page.locator('#menu-conditional')).toBeVisible();
  await page.locator('#menu-conditional [data-cf-submenu="dataBar"]').hover();
  await expect(page.locator('#menu-conditional [data-cf-panel="dataBar"]')).toBeVisible();
  await snapshotPage(page, 'excel-chrome-home-conditional-formatting-menu-1440');
  await page.keyboard.press('Escape');

  await seedTableSelection(page);
  await switchRibbonTab(page, 'insert');
  await snapshotDemo(page, 'excel-chrome-insert-full-1440');
  await page.locator('[data-ribbon-command="formatTableInsert"]').click();
  const createTable = page.getByRole('dialog', { name: 'Create Table' });
  await expect(createTable).toBeVisible();
  await snapshotPage(page, 'excel-chrome-insert-create-table-dialog-1440');
  await createTable.getByRole('button', { name: 'Cancel' }).click();
  await expect(createTable).toBeHidden();
  await page.locator('[data-ribbon-command="pivotTableInsert"]').click();
  const createPivotTable = page.getByRole('dialog', { name: 'Create PivotTable' });
  await expect(createPivotTable).toBeVisible();
  await snapshotPage(page, 'excel-chrome-insert-create-pivottable-dialog-1440');
  await createPivotTable.getByRole('button', { name: 'Cancel' }).click();
  await expect(createPivotTable).toBeHidden();

  await switchRibbonTab(page, 'data');
  await snapshotDemo(page, 'excel-chrome-data-full-1440');
  await switchRibbonTab(page, 'pageLayout');
  await snapshotDemo(page, 'excel-chrome-page-layout-full-1440');
  await page.locator('[data-ribbon-command="pageSetupAdvanced"]').click();
  const pageSetup = page.getByRole('dialog', { name: 'Page Setup' });
  await expect(pageSetup).toBeVisible();
  await snapshotPage(page, 'excel-chrome-page-setup-dialog-1440');
  await pageSetup.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(pageSetup).toBeHidden();

  await switchRibbonTab(page, 'view');
  await snapshotDemo(page, 'excel-chrome-view-full-1440');
  await switchRibbonTab(page, 'file');
  await expect(page.locator('.demo__backstage[role="dialog"]')).toBeVisible();
  await snapshotDemo(page, 'excel-chrome-backstage-info-1440');
  await page
    .locator('.demo__backstage')
    .getByRole('button', { name: 'Print', exact: true })
    .click();
  await snapshotDemo(page, 'excel-chrome-backstage-print-1440');
  await page
    .locator('.demo__backstage')
    .getByRole('button', { name: 'Close', exact: true })
    .click();
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);

  await page.setViewportSize({ width: 1024, height: 768 });
  await switchRibbonTab(page, 'home');
  await setRibbonDisplayMode(page, 'singleLine');
  await snapshotDemo(page, 'excel-chrome-home-single-line-1024');
  await setRibbonDisplayMode(page, 'tabsOnly');
  await snapshotDemo(page, 'excel-chrome-home-tabs-only-1024');
  await setRibbonDisplayMode(page, 'autoHide');
  await page.keyboard.press('Alt');
  await expect(page.locator('.demo__ribbon-shell--autoHidePeek')).toHaveCount(1);
  await snapshotDemo(page, 'excel-chrome-home-auto-hide-peek-1024');

  await page.setViewportSize({ width: 390, height: 844 });
  await setRibbonDisplayMode(page, 'full');
  await expectWorkbookSurfaceChrome(page);
  await snapshotDemo(page, 'excel-chrome-home-full-mobile-390');
}
