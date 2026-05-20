import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

const snapshotDemo = async (page: Page, name: string): Promise<void> => {
  await expect(page.locator('.demo')).toHaveScreenshot(`${name}.png`, {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
  });
};

const switchRibbonTab = async (page: Page, tab: string): Promise<void> => {
  await page.locator(`[data-ribbon-tab="${tab}"]`).click();
  await expect(page.locator(`[data-ribbon-tab="${tab}"][aria-selected="true"]`)).toHaveCount(1);
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
  await switchRibbonTab(page, 'insert');
  await snapshotDemo(page, 'excel-chrome-insert-full-1440');
  await switchRibbonTab(page, 'data');
  await snapshotDemo(page, 'excel-chrome-data-full-1440');
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
