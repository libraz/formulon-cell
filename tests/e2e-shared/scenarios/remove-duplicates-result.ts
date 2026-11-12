import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

type DuplicateStrings = {
  removeDuplicatesDialogTitle: string;
  removeDuplicatesResult: string;
  removeDuplicatesNone: string;
};

/** Data > Remove Duplicates reports what it did in the desktop wording: how many duplicates went
 *  and how many unique values remain, or that none were found. The Mac shows it as a compact
 *  alert. */
export async function runRemoveDuplicatesResultScenario(
  page: Page,
  locale: 'en' | 'ja',
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale, platform });
  await sheet.expectNoStub();
  await sheet.seed({ A1: 'Item', A2: 'Pen', A3: 'Ink', A4: 'Pen', A5: 'Pad', A6: 'Ink' });
  const t = await page.evaluate(
    () =>
      (
        (window as unknown as InstanceWindow).__fcInst.i18n.strings as unknown as {
          ribbonMenu: DuplicateStrings;
        }
      ).ribbonMenu,
  );

  const run = async (range: string) => {
    await page.evaluate(
      (target) => (window as unknown as InstanceWindow).__fcInst.selection.set(target),
      range,
    );
    await selectRibbonTab(page, 'data');
    // The Mac Data tab names the command after its own ribbon id.
    await sheet.clickRibbon(platform === 'mac' ? 'mac.data.removeDuplicates' : 'removeDupes');
    const dialog = page.getByRole('dialog', { name: t.removeDuplicatesDialogTitle });
    await expect(dialog).toBeVisible();
    await dialog.locator('.fc-fmtdlg__btn--primary').click();
    await expect(dialog).toBeHidden();
    const alert = page.locator('[role="alertdialog"]').filter({
      has: page.locator('.fc-tb__dlg__message'),
    });
    await expect(alert).toBeVisible();
    if (platform === 'mac') await expect(alert).toHaveClass(/fc-tb__dlg--alert/);
    return alert;
  };

  let alert = await run('A1:A6');
  await expect(alert.locator('.fc-tb__dlg__message')).toHaveText(
    t.removeDuplicatesResult.replace('{removed}', '2').replace('{remaining}', '3'),
  );
  await alert.getByRole('button').last().click();
  await expect(alert).toBeHidden();
  const column = async () =>
    (await Promise.all(['A2', 'A3', 'A4', 'A5', 'A6'].map((ref) => sheet.readValue(ref)))).map(
      (cell) => (cell.kind === 'text' ? cell.value : null),
    );
  await expect.poll(column).toEqual(['Pen', 'Ink', 'Pad', null, null]);

  // Again over the remaining values only (the two emptied cells would count as a duplicate).
  alert = await run('A1:A4');
  await expect(alert.locator('.fc-tb__dlg__message')).toHaveText(t.removeDuplicatesNone);
  await alert.getByRole('button').last().click();
  await expect(alert).toBeHidden();
  await sheet.expectNoConsoleErrors();
}
