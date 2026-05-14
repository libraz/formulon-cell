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

  // After Enter the cursor is in A2. Step back, then extend to A3.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await sp.shortcut('d');

  // Move to A3 and confirm it picked up "seed".
  await page.keyboard.press('Escape');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  expect(await sp.formulaBarValue()).toBe('seed');
}
