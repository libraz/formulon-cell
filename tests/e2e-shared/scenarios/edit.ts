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

/** E06: editing a precedent recomputes its dependents. Committing the edit
 *  must be enough — no explicit `recalc()` from the host. The formula's
 *  computed value is read off the status-bar aggregates, which summarize the
 *  selection's engine values (the canvas itself isn't queryable). */
export async function runRecalcAfterEditScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();
  const aggs = page.locator('.fc-host__statusbar-aggs').first();

  await sp.focusHost();
  await page.keyboard.type('20');
  await page.keyboard.press('Enter');
  await page.keyboard.type('=A1*2');
  await page.keyboard.press('Enter');

  // Select the formula cell (A2) so the aggregates describe it alone.
  await page.keyboard.press('ArrowUp');
  await expect(aggs).toContainText('40');

  // Re-edit the precedent. Enter leaves the selection back on A2.
  await sp.focusHost();
  await page.keyboard.type('25');
  await page.keyboard.press('Enter');

  await expect(aggs).toContainText('50');
}

/** E02: formulas run through the WASM engine. After SUM the formula bar
 *  carries the formula and the active cell's underlying value is the sum. */
export async function runFormulaScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  // Focus once: typeIntoActiveCell jumps back to A1 on every call.
  for (const value of ['1', '2', '3', '=SUM(A1:A3)']) {
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  }

  // Re-select the formula cell (A4) and check the formula bar.
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('=SUM(A1:A3)');
  const total = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as {
      workbook: {
        getValue(addr: { sheet: number; row: number; col: number }): unknown;
      };
    };
    return inst.workbook.getValue({ sheet: 0, row: 3, col: 0 });
  });
  expect(total).toEqual({ kind: 'number', value: 6 });
}
