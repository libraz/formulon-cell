import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const readAutoFilterRange = (page: Page) =>
  page.evaluate(
    () =>
      (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).getAutoFilter()?.range ??
      null,
  );

const seedRegion = (sheet: UserJourneyPage) =>
  sheet.seed({ A1: 'Item', B1: 'Qty', A2: 'Pen', B2: 3, A3: 'Ink', B3: 5, A4: 'Pad', B4: 2 });

/** Ctrl+Shift+L turns AutoFilter on over the region around the active cell and off again, as the
 *  ribbon's Filter toggle does. */
export async function runFilterShortcutScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);
  await sheet.goTo('B3');

  await sheet.shortcut('Shift+L');
  await expect
    .poll(() => readAutoFilterRange(page))
    .toEqual({ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 });

  await sheet.shortcut('Shift+L');
  await expect.poll(() => readAutoFilterRange(page)).toBeNull();
  await sheet.expectNoConsoleErrors();
}

/** On a protected sheet the Page Layout Background command is unavailable and Ctrl+Shift+L does not
 *  turn AutoFilter on; both come back once the sheet is unprotected. */
export async function runProtectedSheetFilterAndBackgroundScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en' });
  await sheet.expectNoStub();
  await seedRegion(sheet);
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).protect(),
  );

  await selectRibbonTab(page, 'pageLayout');
  const background = page.locator('[data-ribbon-command="sheetBackground"]').first();
  await expect(background).toBeDisabled();

  await sheet.goTo('B3');
  await sheet.shortcut('Shift+L');
  await expect.poll(() => readAutoFilterRange(page)).toBeNull();

  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook.sheet(0).unprotect(),
  );
  await expect(background).toBeEnabled();
  await sheet.goTo('B3');
  await sheet.shortcut('Shift+L');
  await expect
    .poll(() => readAutoFilterRange(page))
    .toEqual({ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 });
  await sheet.expectNoConsoleErrors();
}
