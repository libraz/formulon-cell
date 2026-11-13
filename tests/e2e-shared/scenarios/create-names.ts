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

/** Each defined name with what it refers to, the leading `=` dropped. */
const definedNameRefs = (page: Page) =>
  page.evaluate(() =>
    Object.fromEntries(
      (window as unknown as InstanceWindow).__fcInst.workbook
        .listDefinedNames()
        .map((name) => [name.name, name.formula.replace(/^=/, '')]),
    ),
  );

const selectRange = (page: Page, range: string) =>
  page.evaluate(
    (target) => (window as unknown as InstanceWindow).__fcInst.selection.set(target),
    range,
  );

/** Create Names from Selection: Top row starts checked over a text header row and Left column
 *  does not over numbers; OK names each column after its header with a sheet-qualified range. A
 *  second run asks once per name the workbook already holds, and No keeps that name's old range. */
export async function runCreateNamesScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Qty', B1: 'Price', A2: 3, B2: 10, A3: 5, B3: 20 });
  await selectRange(page, 'A1:B3');

  const dialog = await openCreateNames(page, sheet, platform);
  await expect(dialog.locator('input[data-fc-edge="top"]')).toBeChecked();
  await expect(dialog.locator('input[data-fc-edge="left"]')).not.toBeChecked();
  await expect(dialog.locator('input[data-fc-edge="bottom"]')).not.toBeChecked();
  await expect(dialog.locator('input[data-fc-edge="right"]')).not.toBeChecked();
  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();
  await expect
    .poll(() => definedNameRefs(page))
    .toEqual({
      Price: 'Sheet1!$B$2:$B$3',
      Qty: 'Sheet1!$A$2:$A$3',
    });

  await selectRange(page, 'A1:B2');
  await (await openCreateNames(page, sheet, platform)).locator('.fc-fmtdlg__btn--primary').click();
  const t = await page.evaluate(
    () =>
      (
        (window as unknown as InstanceWindow).__fcInst.i18n.strings as unknown as {
          ribbonMenu: { createNamesTitle: string; createNamesReplace: string };
        }
      ).ribbonMenu,
  );
  // Yes replaces Qty; No keeps Price as it was.
  for (const [name, answer] of [
    ['Qty', 'Yes'],
    ['Price', 'No'],
  ] as const) {
    const prompt = page
      .locator('[role="alertdialog"]')
      .filter({ hasText: t.createNamesReplace.replace('{name}', name) });
    await expect(prompt).toBeVisible();
    // The Mac asks in a compact alert with no title; elsewhere the prompt is titled Create Names.
    if (platform === 'mac') await expect(prompt).toHaveClass(/fc-tb__dlg--alert/);
    else await expect(page.getByRole('alertdialog', { name: t.createNamesTitle })).toBeVisible();
    await prompt.getByRole('button', { name: answer, exact: true }).click();
    await expect(prompt).toBeHidden();
  }
  await expect
    .poll(() => definedNameRefs(page))
    .toEqual({
      Price: 'Sheet1!$B$2:$B$3',
      Qty: 'Sheet1!$A$2',
    });
  await sheet.expectNoConsoleErrors();
}

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
