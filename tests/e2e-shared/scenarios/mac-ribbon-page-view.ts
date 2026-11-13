import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacScaleAndSheetViewScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  const read = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      const state = store.getState();
      return {
        setup: state.pageSetup.setupBySheet.get(0),
        views: state.sheetViews.views,
        freezeRows: state.layout.freezeRows,
      };
    });
  await selectRibbonTab(page, 'pageLayout');
  for (const [id, value] of [
    ['scaleWidth', '1'],
    ['scaleHeight', '2'],
  ]) {
    const select = page.locator(`[data-ribbon-select="${id}"]`);
    await select.locator('.fc-tb__rb-dd__btn').click();
    await select.locator(`[role="option"][data-value="${value}"]`).click();
  }
  await expect.poll(async () => (await read()).setup).toMatchObject({ fitWidth: 1, fitHeight: 2 });
  await selectRibbonTab(page, 'view');
  await sp.clickRibbon('mac.view.freeze');
  await page
    .locator('#menu-mac-view-freeze [data-ribbon-command="mac.view.freeze.firstRow"]')
    .click();
  await sp.clickRibbon('mac.view.sheetViewSave');
  await expect.poll(async () => (await read()).views.length).toBe(1);
  const view = (await read()).views[0];
  if (!view) throw new Error('saved sheet view is missing');
  const id = view.id;
  await sp.clickRibbon('mac.view.freeze');
  // While panes are frozen the menu's first item is Unfreeze Panes.
  const unfreeze = page.locator(
    '#menu-mac-view-freeze [data-ribbon-command="mac.view.freeze.panes"]',
  );
  await expect(unfreeze).toHaveText('Unfreeze Panes');
  await unfreeze.click();
  await expect.poll(async () => (await read()).freezeRows).toBe(0);
  const select = page.locator('[data-ribbon-select="sheetViewSelect"]');
  await select.locator('.fc-tb__rb-dd__btn').click();
  await select.locator(`[role="option"][data-value="${id}"]`).click();
  await expect.poll(async () => (await read()).freezeRows).toBe(1);
  await sp.clickRibbon('mac.view.sheetViewDelete');
  await expect.poll(async () => (await read()).views.length).toBe(0);
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(async () => (await read()).views.length).toBe(1);
  await sp.expectNoConsoleErrors();
}

export async function runMacViewShowScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await selectRibbonTab(page, 'view');
  const formulaBar = page.locator('[data-ribbon-command="mac.view.formulaBar"]').first();
  await expect(formulaBar).toHaveAttribute('aria-pressed', 'true');
  await sp.clickRibbon('mac.view.formulaBar');
  await expect(page.locator('.fc-host__formulabar')).toBeHidden();
  await expect(formulaBar).toHaveAttribute('aria-pressed', 'false');
  await sp.clickRibbon('mac.view.formulaBar');
  await expect(page.locator('.fc-host__formulabar')).toBeVisible();
  await expect(formulaBar).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await sp.expectNoConsoleErrors();
}

export async function runMacPageAndCalculationScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  const read = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      // The calc mode has no facade member; read it from the engine handle.
      const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
      return {
        setup: store.getState().pageSetup.setupBySheet.get(0),
        mode: handle.calcMode(),
      };
    });
  await selectRibbonTab(page, 'pageLayout');
  await sp.clickRibbon('mac.page.orientation');
  await page
    .locator('#menu-mac-page-orientation [data-ribbon-command="mac.page.orientation.landscape"]')
    .click();
  await expect.poll(read).toMatchObject({ setup: { orientation: 'landscape' } });
  await sp.clickRibbon('mac.page.margins');
  await page
    .locator('#menu-mac-page-margins [data-ribbon-command="mac.page.margins.narrow"]')
    .click();
  await expect.poll(read).toMatchObject({ setup: { margins: { left: 0.25, right: 0.25 } } });
  await sp.goTo('C4');
  await sp.clickRibbon('mac.page.pageBreaks');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.page.break.insert"]')
    .click();
  await expect
    .poll(read)
    .toMatchObject({ setup: { manualPageBreakRows: [3], manualPageBreakCols: [2] } });
  await sp.clickRibbon('mac.page.pageBreaks');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.page.break.reset"]')
    .click();
  expect((await read()).setup?.manualPageBreakRows ?? []).toEqual([]);
  await selectRibbonTab(page, 'formulas');
  await sp.clickRibbon('mac.formulas.calcOptions');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.formulas.calc.manual"]')
    .click();
  await expect.poll(read).toMatchObject({ mode: 1 });
  await sp.clickRibbon('mac.formulas.calcOptions');
  await page.keyboard.press('Escape');
  await expect(page.locator('.fc-tb__menu--mac:not([hidden])')).toHaveCount(0);
  await expect(page.locator('[data-ribbon-command="mac.formulas.calcOptions"]')).toBeFocused();
  await sp.clickRibbon('mac.formulas.calcOptions');
  await page
    .locator('[role=menu]:not([hidden]) [data-ribbon-command="mac.formulas.calc.auto"]')
    .click();
  await expect.poll(read).toMatchObject({ mode: 0 });
  await sp.expectNoConsoleErrors();
}

export async function runMacPageToggleAndZoomScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await selectRibbonTab(page, 'pageLayout');
  for (const [action, initial] of [
    ['showGridlines', true],
    ['printGridlines', false],
    ['showHeadings', true],
    ['printHeadings', false],
  ] as const) {
    const button = page.locator(
      `[data-ribbon-panel="pageLayout"] [data-ribbon-command="mac.page.${action}"]`,
    );
    await expect(button).toHaveAttribute('aria-pressed', String(initial));
    await button.click();
    await expect(button).toHaveAttribute('aria-pressed', String(!initial));
  }
  await selectRibbonTab(page, 'view');
  const panel = page.locator('[data-ribbon-panel="view"]');
  const views = panel
    .locator('section')
    .filter({ has: page.locator('[data-ribbon-select="sheetViewSelect"]') });
  await expect(views.locator('[data-ribbon-command="mac.view.standard"]')).toHaveCount(0);
  await sp.clickRibbon('mac.view.zoom');
  const dialog = page.getByRole('dialog');
  await expect(dialog).toBeVisible();
  await dialog.locator('input[type="number"]').fill('150');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  const zoom = () =>
    page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.viewport.zoomPercent /
        100,
    );
  await expect.poll(zoom).toBe(1.5);
  await sp.clickRibbon('mac.view.zoom100');
  await expect.poll(zoom).toBe(1);
  await selectRibbonTab(page, 'pageLayout');
  await expect(page.locator('[data-ribbon-command="mac.page.printGridlines"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await sp.expectNoConsoleErrors();
}
