import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** The pending-format slot of the UI state, when the build has one. */
type PendingFormatUi = {
  pendingFormat?: {
    addr: { sheet: number; row: number; col: number };
    format: { numFmt?: unknown };
  } | null;
};

/** Exercise real keyboard events against the language control and grid. */
export async function runLocaleShortcutsScenario(
  page: Page,
  platform: 'mac' | 'default',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const errors = sp.collectConsoleErrors();
  await sp.mount({ locale: 'en', platform });
  await sp.expectNoStub();
  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    inst.update({
      ui: { ...inst.options.ui, features: { ...inst.options.ui?.features, localeSwitcher: true } },
    });
  });
  const select = page.locator('[data-fc-locale-select]');
  await expect(select).toBeVisible();

  const readFormat = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const state = w.__fcProbes['mount/internals']
        .spreadsheetInternals(w.__fcInst)
        .store.getState();
      const addr = state.selection.active;
      const pending = (state.ui as PendingFormatUi).pendingFormat;
      if (
        pending &&
        pending.addr.sheet === addr.sheet &&
        pending.addr.row === addr.row &&
        pending.addr.col === addr.col
      ) {
        return pending.format.numFmt ?? null;
      }
      return state.format.formats.get(`${addr.sheet}:${addr.row}:${addr.col}`)?.numFmt ?? null;
    });

  for (const [locale, symbol] of [
    ['ja', '¥'],
    ['de', '€'],
    ['fr', '€'],
    ['zh', '¥'],
    ['ko', '₩'],
    ['th', '฿'],
    ['en', '$'],
  ]) {
    await select.focus();
    await select.selectOption(locale);
    await expect(select).toBeFocused();
    await sp.focusHost();
    await page.keyboard.press('Control+Shift+4');
    await expect.poll(readFormat).toEqual({ kind: 'currency', decimals: 2, symbol });
  }

  // Repeating the last currency command must also use the new locale.
  await select.selectOption('ja');
  await sp.focusHost();
  await page.keyboard.press('F4');
  await expect.poll(readFormat).toEqual({ kind: 'currency', decimals: 2, symbol: '¥' });

  await select.focus();
  const before = await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const instance = w.__fcInst;
    const state = w.__fcProbes['mount/internals'].spreadsheetInternals(instance).store.getState();
    return {
      active: state.selection.active,
      pending: (state.ui as PendingFormatUi).pendingFormat,
      formats: [...state.format.formats],
      sheets: instance.workbook.listSheets().length,
    };
  });
  await page.keyboard.press('Control+Shift+4');
  await page.keyboard.press('Shift+F11');
  const after = await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const instance = w.__fcInst;
    const state = w.__fcProbes['mount/internals'].spreadsheetInternals(instance).store.getState();
    return {
      active: state.selection.active,
      pending: (state.ui as PendingFormatUi).pendingFormat,
      formats: [...state.format.formats],
      sheets: instance.workbook.listSheets().length,
    };
  });
  expect(after).toEqual(before);
  await expect(select).toBeFocused();
  const statusId = await select.getAttribute('aria-describedby');
  expect(statusId).toBeTruthy();
  await expect(page.locator(`[id="${statusId}"]`)).toHaveAttribute('role', 'status');
  expect(errors.read()).toEqual([]);
}
