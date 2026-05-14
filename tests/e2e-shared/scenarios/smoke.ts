import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** S01: the host mounts and the engine settles into a ready state.
 *  S02: no `console.error` / pageerror fires during a clean mount. */
export async function runSmokeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();

  await sp.mount();

  // Engine must be in real WASM mode — the demo apps inject COOP/COEP.
  await sp.expectNoStub();

  // crossOriginIsolated must hold or pthread WASM would have fallen back.
  expect(await sp.isCrossOriginIsolated()).toBe(true);

  // Allow async paint / observer chains to settle before assertion.
  await page.waitForTimeout(250);

  expect(consoleErrors.read(), 'expected no console errors during clean mount').toEqual([]);
}
