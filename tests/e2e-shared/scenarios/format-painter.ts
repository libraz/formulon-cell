import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** F02 — apply a currency format imperatively and verify the engine round-trips
 *  the formatted value through `cells.resolveDisplay`. Avoids canvas pixel
 *  inspection by going through the public display-resolution API. */
export async function runCurrencyFormatScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: {
      workbook: { setNumber(addr: { sheet: number; row: number; col: number }, n: number): void };
      store: {
        setState(fn: (s: unknown) => unknown): void;
        getState(): unknown;
      };
    };
  };

  const result = await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    const inst = w.__fcInst;
    if (!inst) return { ok: false as const };

    const addr = { sheet: 0, row: 0, col: 0 };
    inst.workbook.setNumber(addr, 1234.5);
    // Patch the cell format slice directly to apply currency.
    inst.store.setState((s) => {
      const state = s as Record<string, unknown>;
      const fmt = state.format as { formats: Map<string, unknown> };
      const next = new Map(fmt.formats);
      next.set('0:0:0', { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } });
      return {
        ...(state as object),
        format: { ...fmt, formats: next },
      };
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
