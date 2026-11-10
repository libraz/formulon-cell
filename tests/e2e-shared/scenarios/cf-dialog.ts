import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** D05 — the CF Rules dialog is reachable via `instance.dialogs.open('cfRules')`.
 *  Verify calling it from a demo app exposes a visible overlay and that
 *  Esc closes it cleanly. */
export async function runCfRulesDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: SpreadsheetInstance;
  };

  await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    w.__fcInst?.dialogs.open('cfRules');
  });

  // The CF rules dialog uses its own `fc-cfrulesdlg` overlay class.
  await expect(page.locator('.fc-cfrulesdlg')).toBeVisible({ timeout: 2_000 });

  await page.keyboard.press('Escape');
  expect(consoleErrors.read()).toEqual([]);
}

/** D05b — opening the conditional-format authoring dialog through the
 *  `instance.dialogs.open('conditional')` API. */
export async function runConditionalDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: SpreadsheetInstance;
  };

  await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    w.__fcInst?.dialogs.open('conditional');
  });

  await page.waitForTimeout(150);
  // The dialog class is `fc-fmtdlg fc-conddlg`.
  await expect(page.locator('.fc-conddlg')).toBeVisible({ timeout: 2000 });

  await page.keyboard.press('Escape');
  expect(consoleErrors.read()).toEqual([]);
}
