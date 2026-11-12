import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** F2 on a focused sheet tab opens the inline rename input; committing with
 *  Enter writes the new name to both the workbook and the tab DOM. */
export async function runSheetRenameScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const tab = page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first();
  await tab.focus();
  await page.keyboard.press('F2');

  const renameInput = page.locator('.fc-host__sheetbar-rename');
  await expect(renameInput).toBeVisible();

  // Select all + replace.
  await renameInput.fill('PlanA');
  await renameInput.press('Enter');

  await expect(renameInput).toBeHidden();
  await expect(page.locator('.fc-host__sheetbar-tab[aria-selected="true"]')).toHaveText('PlanA');
}

/** ESC during inline rename discards the change and restores the prior name. */
export async function runSheetRenameCancelScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const tab = page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first();
  const original = (await tab.textContent()) ?? '';

  await tab.focus();
  await page.keyboard.press('F2');

  const renameInput = page.locator('.fc-host__sheetbar-rename');
  await expect(renameInput).toBeVisible();

  await renameInput.fill('shouldNotPersist');
  await renameInput.press('Escape');

  await expect(renameInput).toBeHidden();
  await expect(page.locator('.fc-host__sheetbar-tab[aria-selected="true"]')).toHaveText(original);
}

/** History, in any case, is a reserved sheet name off the Mac: Enter leaves the rename input open
 *  and marked invalid, and Escape keeps the old name. The Mac profile accepts it. */
export async function runSheetRenameHistoryScenario(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ platform });
  await sp.expectNoStub();

  const active = page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first();
  const original = (await active.textContent()) ?? '';
  const renameInput = page.locator('.fc-host__sheetbar-rename');
  const names = platform === 'mac' ? ['History'] : ['History', 'history', 'HISTORY'];
  for (const name of names) {
    await active.focus();
    await page.keyboard.press('F2');
    await expect(renameInput).toBeVisible();
    await renameInput.fill(name);
    await renameInput.press('Enter');
    if (platform === 'mac') {
      await expect(renameInput).toBeHidden();
      await expect(active).toHaveText(name);
      return;
    }
    await expect(renameInput).toBeVisible();
    await expect(renameInput).toHaveAttribute('aria-invalid', 'true');
    await renameInput.press('Escape');
    await expect(renameInput).toBeHidden();
    await expect(active).toHaveText(original);
  }
}
