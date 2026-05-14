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

  // Seed A1=1, B1=2, A2=3 by walking arrows + Enter.
  await sp.typeIntoActiveCell('1');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('2');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.type('3');
  await page.keyboard.press('Enter');

  // Re-walk and verify the values.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('1');

  await page.keyboard.press('ArrowRight');
  expect(await sp.formulaBarValue()).toBe('2');

  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowLeft');
  expect(await sp.formulaBarValue()).toBe('3');
}
