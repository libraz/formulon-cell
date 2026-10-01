import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** S01: the host mounts and the engine settles into a ready state.
 *  S02: no `console.error` / pageerror fires during a clean mount. */
export async function runSmokeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();

  await sp.mount();

  // The default serial engine must mount without isolation or stub fallback.
  await sp.expectNoStub();

  // These demos serve no COOP/COEP headers.
  expect(await sp.isCrossOriginIsolated()).toBe(false);

  // Allow async paint / observer chains to settle before assertion.
  await page.waitForTimeout(250);

  expect(consoleErrors.read(), 'expected no console errors during clean mount').toEqual([]);
}
