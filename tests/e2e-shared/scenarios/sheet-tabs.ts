import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** T01 — adding a sheet via the `+` button in the sheetbar increases the
 *  tab count and makes the new tab the selected one.
 *
 *  Uses only core chrome (`.fc-host__sheetbar-*`), so it runs identically
 *  across the three demo apps. */
export async function runSheetTabsScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const tabs = page.locator('.fc-host__sheetbar-tab');
  const initial = await tabs.count();
  expect(initial, 'demo apps boot with at least one sheet').toBeGreaterThanOrEqual(1);

  await page.locator('.fc-host__sheetbar-add').click();

  // The new tab gets aria-selected="true". Wait until the count changes so we
  // don't race the re-render in `update()` (sheet-tabs-controller.ts:221).
  await expect(tabs).toHaveCount(initial + 1);

  const selected = page.locator('.fc-host__sheetbar-tab[aria-selected="true"]');
  await expect(selected).toHaveCount(1);
}

/** T02 — clicking a sibling tab switches the active sheet. */
export async function runSheetTabSwitchScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Make sure we have at least two sheets.
  const addBtn = page.locator('.fc-host__sheetbar-add');
  const tabs = page.locator('.fc-host__sheetbar-tab');
  if ((await tabs.count()) < 2) {
    await addBtn.click();
    await expect(tabs).toHaveCount(2);
  }

  // Click the first tab and verify selection moves.
  await tabs.first().click();
  await expect(tabs.first()).toHaveAttribute('aria-selected', 'true');
}
