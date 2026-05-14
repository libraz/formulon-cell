import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** U01 — type → Mod+Z reverts; Mod+Shift+Z (Y on non-mac) re-applies.
 *  We test the cross-OS branch via shortcut() and assert via the formula bar. */
export async function runUndoRedoScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('gamma');

  // Back to A1, verify the commit.
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('gamma');

  // Undo — A1 should become empty.
  await sp.shortcut('z');
  expect(await sp.formulaBarValue()).toBe('');

  // Redo — A1 returns. On macOS this is Shift+Cmd+Z; elsewhere Ctrl+Y.
  const isMac = await page.evaluate(() => navigator.platform.toLowerCase().includes('mac'));
  if (isMac) await page.keyboard.press('Meta+Shift+Z');
  else await page.keyboard.press('Control+Y');

  expect(await sp.formulaBarValue()).toBe('gamma');
}
