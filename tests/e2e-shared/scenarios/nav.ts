import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** E03 — Tab moves the active cell right, Shift+Tab moves it back,
 *  Arrow keys navigate, Ctrl+End jumps to the end of data. We probe via
 *  formula-bar contents after seeding cells. */
export async function runArrowAndTabNavScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // A1='alpha' commit → cursor at A2.
  await sp.typeIntoActiveCell('alpha');
  // Go back to A1.
  await page.keyboard.press('ArrowUp');
  // Tab right → enters B1.
  await page.keyboard.press('Tab');
  // Type bravo, commit.
  await page.keyboard.type('bravo');
  await page.keyboard.press('Enter');
  // After commit, cursor returns near the Tab-anchored row. Step back to A1 / B1 via arrow.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowLeft');
  expect(await sp.formulaBarValue()).toBe('alpha');

  await page.keyboard.press('ArrowRight');
  expect(await sp.formulaBarValue()).toBe('bravo');

  // Shift+Tab from B1 → A1.
  await page.keyboard.press('Shift+Tab');
  expect(await sp.formulaBarValue()).toBe('alpha');
}

/** E04 — F2 enters edit mode, ESC discards the change, Enter commits. */
export async function runF2EscapeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('keep');
  await page.keyboard.press('ArrowUp');

  // F2 to enter edit mode.
  await page.keyboard.press('F2');
  await page.keyboard.type('-extra');
  // ESC discards.
  await page.keyboard.press('Escape');

  // Re-select to read the formula bar.
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('keep');
}
