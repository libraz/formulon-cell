import { expect, type Page } from '@playwright/test';

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
