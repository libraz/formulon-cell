import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Opens Formulas > Create from Selection: the Names split menu on default chrome, the plain
 *  command on the Mac. */
const openCreateNames = async (page: Page, sheet: UserJourneyPage, platform: 'default' | 'mac') => {
  await selectRibbonTab(page, 'formulas');
  if (platform === 'mac') {
    await sheet.clickRibbon('mac.formulas.createFromSelection');
  } else {
    // Names opens the Name Manager; its chevron opens the Defined Names menu.
    await page
      .locator('[data-ribbon-command="namedRanges"]')
      .first()
      .locator('.fc-tb__rb-split-chevron')
      .click();
    await page.locator('[data-defined-name-action="create-from-selection"]:visible').click();
  }
  const dialog = page.locator('.fc-createnames');
  await expect(dialog).toBeVisible();
  return dialog;
};

const definedNames = (page: Page) =>
  page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.workbook
      .listDefinedNames()
      .map((name) => name.name)
      .sort(),
  );

const selectRange = (page: Page, range: string) =>
  page.evaluate(
    (target) => (window as unknown as InstanceWindow).__fcInst.selection.set(target),
    range,
  );

/** Create Names from Selection over two whole rows refuses with the desktop's "Selection is too
 *  large." alert and defines no names; the Mac shows it as a compact alert. */
export async function runCreateNamesTooLargeScenario(
  page: Page,
  locale: 'en' | 'ja',
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale, platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Qty', A2: 3 });
  await selectRange(page, 'A1:XFD2');
  const tooLarge = await page.evaluate(
    () =>
      (
        (window as unknown as InstanceWindow).__fcInst.i18n.strings as unknown as {
          commandRefusal: { tooLarge: string };
        }
      ).commandRefusal.tooLarge,
  );

  const dialog = await openCreateNames(page, sheet, platform);
  const ok = dialog.locator('.fc-fmtdlg__btn--primary');
  for (const edge of ['top', 'left', 'bottom', 'right'])
    await dialog.locator(`input[data-fc-edge="${edge}"]`).uncheck();
  // With no edge checked there is nothing to create, so OK is greyed.
  await expect(ok).toBeDisabled();
  await dialog.locator('input[data-fc-edge="top"]').check();
  await expect(ok).toBeEnabled();
  await ok.click();

  const alert = page.locator('[role="alertdialog"]').filter({ hasText: tooLarge });
  await expect(alert).toBeVisible();
  if (platform === 'mac') await expect(alert).toHaveClass(/fc-tb__dlg--alert/);
  await alert.getByRole('button').last().click();
  await expect(alert).toBeHidden();
  expect(await definedNames(page)).toEqual([]);
  await sheet.expectNoConsoleErrors();
}
