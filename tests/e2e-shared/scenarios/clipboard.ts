import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** C01 — internal copy/paste round-trip via Mod+C/V keyboard shortcuts.
 *  Canvas content isn't queryable, so we re-select each cell and read the
 *  formula bar to confirm the value landed. */
export async function runCopyPasteScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Seed A1=alpha; cursor advances to A2 on Enter.
  await sp.typeIntoActiveCell('alpha');
  // Step back to A1, copy.
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('c');

  // Navigate to A3 and paste. Active cell after paste is the paste anchor.
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await sp.shortcut('v');

  // Mod+V routes through navigator.clipboard.readText() which is async, so
  // poll the formula bar instead of asserting immediately.
  await expect.poll(() => sp.formulaBarValue(), { timeout: 2_000 }).toBe('alpha');
}

/** C02 — Mod+X cut → paste removes the source.
 *  Source becomes empty after paste; destination holds the value. */
export async function runCutPasteScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('beta');
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('x');

  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await sp.shortcut('v');

  // Destination has the value (paste is async — see C01).
  await expect.poll(() => sp.formulaBarValue(), { timeout: 2_000 }).toBe('beta');

  // ...and the source (A1) is empty.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('');
}
