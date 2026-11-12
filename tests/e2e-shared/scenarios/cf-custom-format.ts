import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

const visibleTabs = (dialog: Locator) =>
  dialog
    .locator('.fc-fmtdlg__tabs [role="tab"]:visible')
    .evaluateAll((els) => els.map((el) => el.textContent ?? ''));

/** New Rule > Custom Format...: on the Mac, Format Cells keeps Number but opens on Font, greys the
 *  font name, size, underline and super/subscript, shows strikethrough mixed and greys Clear until a
 *  change; elsewhere it is the full form on Number. */
export async function runConditionalCustomFormatScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja', platform });
  await sheet.expectNoStub();
  await sheet.goTo('A1');
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.dialogs.open('conditional'),
  );
  const rule = page.locator('.fc-conddlg');
  await expect(rule).toBeVisible();
  const preset = rule
    .locator('select:visible')
    .filter({ has: page.locator('option[value="custom"]') })
    .first();
  await preset.selectOption('custom');

  const dialog = page
    .locator('.fc-fmtdlg__panel:visible:not(.fc-conddlg__panel)')
    .filter({ has: page.locator('[role="tab"][data-fc-tab="font"]') })
    .last();
  await expect(dialog).toBeVisible();
  const selected = dialog.locator('.fc-fmtdlg__tabs [role="tab"][aria-selected="true"]');
  const clear = dialog.locator('.fc-fmtdlg__btn--clear');
  const family = dialog.locator('input[data-fc-input="family"]');
  const underline = dialog.locator('select[data-fc-input="underline"]');

  if (platform === 'default') {
    await expect(selected).toHaveText('表示形式');
    await expect(clear).toBeHidden();
    await dialog.locator('[role="tab"][data-fc-tab="font"]').click();
    await expect(family).toBeEnabled();
    await expect(underline).toBeEnabled();
  } else {
    expect(await visibleTabs(dialog)).toEqual(['表示形式', 'フォント', '罫線', '塗りつぶし']);
    await expect(selected).toHaveText('フォント');
    await expect(family).toBeDisabled();
    await expect(dialog.locator('input[data-fc-input="size"]')).toBeDisabled();
    await expect(underline).toBeDisabled();
    await expect(dialog.locator('input[data-fc-check="superscript"]')).toBeDisabled();
    await expect(dialog.locator('input[data-fc-check="subscript"]')).toBeDisabled();
    const strike = dialog.locator('input[data-fc-check="strike"]');
    expect(await strike.evaluate((el) => (el as HTMLInputElement).indeterminate)).toBe(true);
    await expect(clear).toBeVisible();
    await expect(clear).toBeDisabled();
    // The Mac form picks italic from the style list.
    await dialog.locator('[data-fc-font-style="italic"]').click();
    await expect(dialog.locator('[data-fc-font-style="italic"]')).toHaveAttribute(
      'aria-selected',
      'true',
    );
    await expect(clear).toBeEnabled();
    await clear.click();
    await expect(dialog.locator('[data-fc-font-style="italic"]')).not.toHaveAttribute(
      'aria-selected',
      'true',
    );
    await expect(clear).toBeDisabled();
  }

  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();
  await expect(rule).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(rule).toBeHidden();
  await sheet.expectNoConsoleErrors();
}
