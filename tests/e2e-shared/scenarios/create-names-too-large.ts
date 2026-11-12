import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Formulas > Create from Selection > Top Row over two whole rows refuses with the desktop's
 *  "Selection is too large." alert and defines no names; the Mac shows it as a compact alert. */
export async function runCreateNamesTooLargeScenario(
  page: Page,
  locale: 'en' | 'ja',
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale, platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Qty', A2: 3 });
  const tooLarge = await page.evaluate(() => {
    const inst = (window as unknown as InstanceWindow).__fcInst;
    inst.selection.set('A1:XFD2');
    return (inst.i18n.strings as unknown as { commandRefusal: { tooLarge: string } }).commandRefusal
      .tooLarge;
  });

  await selectRibbonTab(page, 'formulas');
  if (platform === 'mac') {
    await sheet.clickRibbon('mac.formulas.createFromSelection');
    await page
      .locator(
        '#menu-mac-formulas-create-names [data-ribbon-command="mac.formulas.createNames.topRow"]',
      )
      .click();
  } else {
    // Names opens the Name Manager; its chevron opens the Defined Names menu.
    await page
      .locator('[data-ribbon-command="namedRanges"]')
      .first()
      .locator('.fc-tb__rb-split-chevron')
      .click();
    await page.locator('[data-defined-name-action="create-top-row"]:visible').click();
  }

  const alert = page.locator('[role="alertdialog"]').filter({ hasText: tooLarge });
  await expect(alert).toBeVisible();
  if (platform === 'mac') await expect(alert).toHaveClass(/fc-tb__dlg--alert/);
  await alert.getByRole('button').last().click();
  await expect(alert).toBeHidden();
  expect(
    await page.evaluate(
      () => (window as unknown as InstanceWindow).__fcInst.workbook.listDefinedNames().length,
    ),
  ).toBe(0);
  await sheet.expectNoConsoleErrors();
}
