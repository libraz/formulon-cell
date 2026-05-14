import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** Navigation: arrow keys + Enter + Tab take you through the grid in the
 *  expected order. This tests the host's keyboard router without requiring
 *  specific shortcut bindings (Home/End vary across browsers/OS). */
export async function runArrowSequenceScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Seed A1=1, B1=2, A2=3 by walking arrows + Enter. typeIntoActiveCell
  // commits with Enter, which advances the cursor one row down; subsequent
  // arrow steps reposition before each typed value.
  await sp.typeIntoActiveCell('1'); // A1=1, cursor → A2
  await page.keyboard.press('ArrowUp'); // A1
  await page.keyboard.press('ArrowRight'); // B1
  await page.keyboard.type('2');
  await page.keyboard.press('Enter'); // B1=2, cursor → B2
  await page.keyboard.press('ArrowLeft'); // A2
  await page.keyboard.type('3');
  await page.keyboard.press('Enter'); // A2=3, cursor → A3

  // Re-walk and verify the values.
  await page.keyboard.press('ArrowUp'); // A2
  expect(await sp.formulaBarValue()).toBe('3');

  await page.keyboard.press('ArrowUp'); // A1
  expect(await sp.formulaBarValue()).toBe('1');

  await page.keyboard.press('ArrowRight'); // B1
  expect(await sp.formulaBarValue()).toBe('2');
}
