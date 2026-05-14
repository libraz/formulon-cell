import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** N01 — `=1/0` is preserved verbatim in the formula bar; the engine resolves
 *  the cell to `#DIV/0!` (visible only on the canvas, so we don't assert text).
 *  The acceptance bar of this test is "no console error fires" — the engine
 *  must not throw when it produces an error value. */
export async function runDivByZeroScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('=1/0');
  await page.keyboard.press('ArrowUp');

  expect(await sp.formulaBarValue()).toBe('=1/0');

  // Allow the engine to evaluate before draining errors.
  await page.waitForTimeout(150);
  expect(consoleErrors.read(), 'engine should not log when producing #DIV/0!').toEqual([]);
}

/** N02 — `=NOTAFN()` is an unknown function; the engine must resolve it to a
 *  `#NAME?` value rather than throwing. Same formula-bar persistence check. */
export async function runUnknownFunctionScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('=NOTAFN()');
  await page.keyboard.press('ArrowUp');

  expect(await sp.formulaBarValue()).toBe('=NOTAFN()');
  await page.waitForTimeout(150);
  expect(consoleErrors.read(), 'engine should not log on unknown function').toEqual([]);
}
