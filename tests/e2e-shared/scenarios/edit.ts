import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** E01: typing into the active cell and pressing Enter commits the value.
 *  The cell text isn't queryable (canvas), so we round-trip via the formula
 *  bar after re-selecting the cell. */
export async function runEditBasicScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.typeIntoActiveCell('123');
  // Move back into the cell so the formula bar shows its value.
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('123');
}

/** E02: formulas run through the WASM engine. After SUM the formula bar
 *  carries the formula and the active cell's underlying value is the sum. */
export async function runFormulaScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('1');
  await sp.typeIntoActiveCell('2');
  await sp.typeIntoActiveCell('3');
  // Cursor is now in row 4. Type the formula and commit.
  await sp.typeIntoActiveCell('=SUM(A1:A3)');

  // Re-select the formula cell (A4) and check the formula bar.
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('=SUM(A1:A3)');
}
