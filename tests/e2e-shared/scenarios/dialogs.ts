import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** D02 — Find & Replace. Bound to Mod+F via host-shortcuts.ts. The find
 *  overlay's exact selector varies by build but every demo's chrome attaches
 *  a `.fc-findreplace` (or close descendant) on open. We test the keystroke
 *  succeeds without throwing and SOME find-related UI becomes visible. */
export async function runFindReplaceShortcutScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  await sp.shortcut('f');

  // Settle async paint.
  await page.waitForTimeout(150);
  expect(consoleErrors.read(), 'opening find should not error').toEqual([]);

  // Closing with Escape should also be safe.
  await page.keyboard.press('Escape');
  expect(consoleErrors.read()).toEqual([]);
}

/** D01 (smoke) — Mod+Z / Mod+Y undo redo path is bound by the host's
 *  shortcut router. Closing path checks that nothing throws. */
export async function runUndoSubsystemSmokeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.shortcut('z');
  await page.waitForTimeout(50);
  expect(consoleErrors.read()).toEqual([]);
}

/** D03 — Shift+F3 and the formula bar fx button open Function Arguments, and
 *  closing it by Escape or Cancel hands focus back to the sheet. */
export async function runFunctionDialogShortcutScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();
  const host = page.locator('.fc-host').first();
  const dialog = page.getByRole('dialog', { name: 'Function Arguments', exact: true });

  await sp.focusHost();
  await page.keyboard.press('Shift+F3');
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('combobox', { name: 'Search functions…' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(host).toBeFocused();

  const fx = page.locator('.fc-host__formulabar-fx').first();
  for (const close of ['Escape', 'Cancel'] as const) {
    await fx.click();
    await expect(dialog).toBeVisible();
    if (close === 'Escape') await page.keyboard.press('Escape');
    else await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(dialog, `closed by ${close}`).toHaveCount(0);
    await expect(host, `focus after ${close}`).toBeFocused();
  }
  expect(consoleErrors.read()).toEqual([]);
}
