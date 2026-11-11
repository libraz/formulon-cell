import { expect, type Page } from '@playwright/test';

import type { ProtectOptions } from '../../../packages/formulon-cell/src/api/sheet-format.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const protectSheet = (page: Page, allowEditObjects: boolean) =>
  page.evaluate(async (allow) => {
    const sheet = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await sheet.unprotect();
    await sheet.protect({ allowEditObjects: allow });
  }, allowEditObjects);

/** On a protected sheet, inserting a chart or a shape needs the Edit objects option; the cells under
 *  the object need not be unlocked. */
export async function runProtectedSheetObjectScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Item', B1: 'Qty', A2: 'Pen', B2: 3, A3: 'Ink', B3: 5 });
  await sheet.goTo('A1:B3');
  await selectRibbonTab(page, 'insert');
  const chart = page.locator('[data-ribbon-command="chartInsert"]').first();
  const shapes = page.locator('[data-ribbon-command="shapesInsert"]').first();

  await protectSheet(page, false);
  await expect(chart).toBeDisabled();
  await expect(shapes).toBeDisabled();

  await protectSheet(page, true);
  await expect(chart).toBeEnabled();
  await expect(shapes).toBeEnabled();
  await sheet.expectNoConsoleErrors();
}

const protectWith = (page: Page, options: ProtectOptions) =>
  page.evaluate(async (opts) => {
    const sheet = (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0);
    await sheet.unprotect();
    await sheet.protect(opts);
  }, options);

/** Values whose slicer chips read as selected. */
const selectedChips = (page: Page) =>
  page
    .locator('.fc-slicer__chip.fc-slicer__chip--on')
    .evaluateAll((chips) => chips.map((chip) => (chip as HTMLElement).dataset.fcValue));

/** A slicer is a locked drawing: on a protected sheet it is inert without Edit objects (no filter,
 *  no alert). With Edit objects but without Use AutoFilter a click raises the protection alert and
 *  filters nothing; with both it filters. */
export async function runProtectedSheetSlicerScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Region', B1: 'Sales', A2: 'East', B2: 3, A3: 'West', B3: 5 });
  await page.evaluate(() => {
    const instance = (window as unknown as InstanceWindow).__fcInst;
    instance.update({ ui: { ...instance.options.ui, features: { slicer: true } } });
    instance.workbook.sheet(0).addTable({ range: 'A1:B3', name: 'Sales' });
    const slicers = instance.extensions.get('slicer') as unknown as {
      addSlicer(input: { tableName: string; column: string; x?: number; y?: number }): unknown;
    } | null;
    if (!slicers) throw new Error('the slicer extension is not mounted');
    slicers.addSlicer({ tableName: 'Sales', column: 'Region', x: 240, y: 24 });
    instance.history.clear();
  });
  const slicer = page.locator('.fc-slicer').first();
  await expect(slicer).toBeVisible();
  const east = slicer.locator('.fc-slicer__chip[data-fc-value="East"]');
  const alert = page.locator('.fc-valdlg:visible');
  expect(await selectedChips(page)).toEqual(['East', 'West']);

  await protectWith(page, {});
  await east.click();
  await page.waitForTimeout(300);
  await expect(alert).toHaveCount(0);
  expect(await selectedChips(page)).toEqual(['East', 'West']);
  await expect(slicer.locator('.fc-slicer__clear')).toBeDisabled();

  await protectWith(page, { allowEditObjects: true });
  await east.click();
  await expect(alert).toBeVisible();
  await page.keyboard.press('Enter');
  await expect(alert).toHaveCount(0);
  expect(await selectedChips(page)).toEqual(['East', 'West']);

  await protectWith(page, { allowEditObjects: true, allowAutoFilter: true });
  await east.click();
  await expect.poll(() => selectedChips(page)).toEqual(['East']);
  await expect(alert).toHaveCount(0);
  await sheet.expectNoConsoleErrors();
}
