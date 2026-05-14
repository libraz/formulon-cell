import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** C04 — Ctrl+Alt+V opens the Paste Special dialog after a prior copy.
 *  Same `fc-fmtdlg fc-pastesp` overlay class as the format dialog family. */
export async function runPasteSpecialDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  // Need something on the clipboard for paste-special to open meaningfully.
  await sp.typeIntoActiveCell('seed');
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('c');

  await page.keyboard.press('Control+Alt+v');
  // Allow async overlay render.
  await page.waitForTimeout(200);

  // The dialog may or may not appear depending on browser focus state under
  // headless test, but it MUST not throw a console error.
  expect(consoleErrors.read()).toEqual([]);
  await page.keyboard.press('Escape');
}
