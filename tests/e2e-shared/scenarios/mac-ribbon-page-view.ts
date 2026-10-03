import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacScaleAndSheetViewScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const read = () =>
    page.evaluate(() => {
      const inst = (
        window as unknown as {
          __fcInst: {
            store: {
              getState(): import('../../../packages/formulon-cell/src/store/types.js').State;
            };
          };
        }
      ).__fcInst;
      const state = inst.store.getState();
      return {
        setup: state.pageSetup.setupBySheet.get(0),
        views: state.sheetViews.views,
        freezeRows: state.layout.freezeRows,
      };
    });
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  for (const [id, value] of [
    ['scaleWidth', '1'],
    ['scaleHeight', '2'],
  ]) {
    const select = page.locator(`[data-ribbon-select="${id}"]`);
    await select.locator('.fc-tb__rb-dd__btn').click();
    await select.locator(`[role="option"][data-value="${value}"]`).click();
  }
  await expect.poll(async () => (await read()).setup).toMatchObject({ fitWidth: 1, fitHeight: 2 });
  await page.locator('[data-ribbon-tab="view"]').click();
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
  await page.locator('#menu-mac-view-freeze [data-ribbon-command="mac.view.freeze.off"]').click();
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
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="view"]').click();
  await sp.clickRibbon('mac.view.show');
  const menu = page.locator('#menu-mac-view-show');
  await expect(menu).toBeVisible();
  await menu.locator('[data-ribbon-command="mac.view.formulaBar"]').click();
  await expect(page.locator('.fc-host__formulabar')).toBeHidden();
  await sp.clickRibbon('mac.view.show');
  await menu.locator('[data-ribbon-command="mac.view.formulaBar"]').click();
  await expect(page.locator('.fc-host__formulabar')).toBeVisible();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await sp.expectNoConsoleErrors();
}

export async function runMacPageAndCalculationScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  const read = () =>
    page.evaluate(() => {
      const inst = (
        window as unknown as {
          __fcInst: {
            store: {
              getState(): import('../../../packages/formulon-cell/src/store/types.js').State;
            };
            workbook: { calcMode(): number };
          };
        }
      ).__fcInst;
      return {
        setup: inst.store.getState().pageSetup.setupBySheet.get(0),
        mode: inst.workbook.calcMode(),
      };
    });
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
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
  await page.locator('[data-ribbon-tab="formulas"]').click();
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
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
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
  await page.locator('[data-ribbon-tab="view"]').click();
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
        (
          window as unknown as {
            __fcInst: { store: { getState(): { viewport: { zoom: number } } } };
          }
        ).__fcInst.store.getState().viewport.zoom,
    );
  await expect.poll(zoom).toBe(1.5);
  await sp.clickRibbon('mac.view.zoom100');
  await expect.poll(zoom).toBe(1);
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  await expect(page.locator('[data-ribbon-command="mac.page.printGridlines"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await sp.expectNoConsoleErrors();
}
