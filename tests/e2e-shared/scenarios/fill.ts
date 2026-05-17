import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Fill-down (Mod+D) — type a value in A1, extend the selection down with
 *  Shift+ArrowDown, then Mod+D. A3 should hold the same value. */
export async function runFillDownScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('seed');

  // After Enter the cursor is in A2. Step back, then extend to A3 so the
  // selection covers A1:A3 with active at A3.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await sp.shortcut('d');

  // Active cell is A3 after the extend; if the fill worked, the formula bar
  // already reflects the seeded value.
  expect(await sp.formulaBarValue()).toBe('seed');

  // Verify A2 also picked it up (mid-range, not just the anchor copy).
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('seed');
}

/** Ribbon Fill > Down — exercises the dropdown-menu click delegator core
 *  auto-wires when the host passes `dynamicDropdowns: true`. Seeds A1,
 *  selects A1:A3, opens the Fill dropdown via the ribbon button, then
 *  clicks the "down" menu item. A3 must hold the seeded value. */
export async function runRibbonFillDownScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('ribbon-fill');

  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');

  await page.locator('[data-ribbon-command="fillHome"]').click();
  await page.locator('#menu-fill [data-fill="down"]').click();

  // Active cell stays at A3 (selection anchor unchanged). Confirm via the
  // formula bar to avoid querying canvas content directly.
  await expect.poll(() => sp.formulaBarValue(), { timeout: 2_000 }).toBe('ribbon-fill');

  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('ribbon-fill');
}
