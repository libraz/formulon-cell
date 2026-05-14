import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** D04 — Data Validation lives in the format dialog under the `more` tab.
 *  Open the dialog via Mod+1, switch to the "More" tab, and verify the
 *  validation kind selector is present and accepts a value. */
export async function runValidationDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  await sp.shortcut('1');
  await expect(page.locator('[class="fc-fmtdlg"]')).toBeVisible({ timeout: 2000 });

  // Click the "More" tab.
  await page.locator('button.fc-fmtdlg__tab[data-fc-tab="more"]').click();
  const morePanel = page.locator('[role="tabpanel"][data-fc-tab="more"]');
  await expect(morePanel).toBeVisible();

  // Switch validation kind to "whole" — the bound-pair rows appear.
  const kindSelect = morePanel.locator('select').first();
  await kindSelect.selectOption('whole');
  await expect(kindSelect).toHaveValue('whole');

  await page.keyboard.press('Escape');
}
