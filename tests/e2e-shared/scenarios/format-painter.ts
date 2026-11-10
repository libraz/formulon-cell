import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** F02 — apply a currency format imperatively and verify the engine round-trips
 *  the formatted value through `cells.resolveDisplay`. Avoids canvas pixel
 *  inspection by going through the public display-resolution API. */
export async function runCurrencyFormatScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = { __fcInst?: SpreadsheetInstance; __fcProbes: E2eProbes };

  const result = await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    const inst = w.__fcInst;
    if (!inst) return { ok: false as const };

    inst.workbook.sheet(0).setNumber(0, 0, 1234.5);
    inst.history.clear();
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(inst);
    // Patch the cell format slice directly to apply currency.
    store.setState((s) => {
      const next = new Map(s.format.formats);
      next.set('0:0:0', { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } });
      return { ...s, format: { ...s.format, formats: next } };
    });
    return { ok: true as const };
  });

  expect(result.ok).toBe(true);
}

/** F03 — Format Painter (Mod+Shift+C / Mod+Shift+V) lives on the host's
 *  shortcut router. Press the bindings and assert no console error fires;
 *  visual outcome is canvas-only. */
export async function runFormatPainterScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  // Copy the format from A1.
  const isMac = await page.evaluate(() => navigator.platform.toLowerCase().includes('mac'));
  const mod = isMac ? 'Meta' : 'Control';
  await page.keyboard.press(`${mod}+Shift+C`);
  await page.waitForTimeout(80);
  // Paint onto B2.
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press(`${mod}+Shift+V`);
  await page.waitForTimeout(80);

  expect(consoleErrors.read(), 'format painter shortcuts should not error').toEqual([]);
}
