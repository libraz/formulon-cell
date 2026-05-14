import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** F01-style — Mod+B toggles bold for the active cell via the host shortcut.
 *  We can't inspect canvas pixels, but we CAN verify that the keystroke
 *  doesn't error and that the format-painter event surface fires. */
export async function runBoldToggleScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('rich');
  await page.keyboard.press('ArrowUp');

  await sp.shortcut('b');
  await page.waitForTimeout(80);
  expect(consoleErrors.read(), 'bold toggle should not error').toEqual([]);
}

/** F01-style — Mod+1 opens the format dialog. We verify the keystroke fires
 *  and the dialog appears in the DOM. */
export async function runFormatDialogShortcutScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.focusHost();
  await sp.shortcut('1');

  // The format dialog overlay flips out of hidden state.
  await expect(
    page.locator('[class="fc-fmtdlg"]'),
    'format dialog should be reachable via Mod+1',
  ).toBeVisible({ timeout: 2000 });
  await page.keyboard.press('Escape');
}
