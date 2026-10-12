import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { expect, type Page } from '@playwright/test';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

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
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.setFeatures({
      localeSwitcher: true,
    });
  });
  const select = page.locator('[data-fc-locale-select]');
  await expect(select).toBeVisible();

  const readFormat = () =>
    page.evaluate(() => {
      const state = (
        window as unknown as { __fcInst: SpreadsheetInstance }
      ).__fcInst.store.getState();
      const addr = state.selection.active;
      const pending = state.ui.pendingFormat;
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
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const state = instance.store.getState();
    return {
      active: state.selection.active,
      pending: state.ui.pendingFormat,
      formats: [...state.format.formats],
      sheets: instance.workbook.sheetCount,
    };
  });
  await page.keyboard.press('Control+Shift+4');
  await page.keyboard.press('Shift+F11');
  const after = await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const state = instance.store.getState();
    return {
      active: state.selection.active,
      pending: state.ui.pendingFormat,
      formats: [...state.format.formats],
      sheets: instance.workbook.sheetCount,
    };
  });
  expect(after).toEqual(before);
  await expect(select).toBeFocused();
  const statusId = await select.getAttribute('aria-describedby');
  expect(statusId).toBeTruthy();
  await expect(page.locator(`[id="${statusId}"]`)).toHaveAttribute('role', 'status');
  expect(errors.read()).toEqual([]);
}
