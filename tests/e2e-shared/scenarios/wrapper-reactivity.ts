import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Both demo wrappers render a "Selection" card driven by `useSelection`
 *  (React hook / Vue composable). The card text is the active cell ref —
 *  it updates after every store mutation, so it's a direct probe of wrapper
 *  reactivity through the host's store subscription.
 *
 *  Test plan:
 *   1. Open the Options panel (toggle button labelled "Options").
 *   2. Capture the initial selection label.
 *   3. Move the active cell (Arrow keys → store mutation).
 *   4. Assert the label changed.
 *
 *  Playground is excluded — it uses raw `Spreadsheet.mount` and has no
 *  reactive selection mirror.
 */
export async function runWrapperReactivityScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Open the Options panel.
  const optionsBtn = page.getByRole('button', { name: 'Options' });
  if ((await optionsBtn.getAttribute('aria-pressed')) === 'false') {
    await optionsBtn.click();
  }

  const selectionCard = page
    .locator('aside[aria-label="Options panel"] section.demo__card')
    .filter({ has: page.locator('h2', { hasText: 'Selection' }) });
  const label = selectionCard.locator('.demo__mono');

  await expect(label).toBeVisible();
  const before = (await label.textContent())?.trim() ?? '';

  // Drive a store mutation: focus the host, move two cells right + one down.
  await sp.focusHost();
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowDown');

  // The wrapper subscribes to the store, so the label re-renders.
  await expect(label).not.toHaveText(before);
}
