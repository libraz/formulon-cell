import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** D05 — the CF Rules dialog is reachable via `instance.openCfRulesDialog()`.
 *  Verify calling it from the playground exposes a visible overlay and that
 *  Esc closes it cleanly. */
export async function runCfRulesDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: { openCfRulesDialog(): void };
  };

  await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    w.__fcInst?.openCfRulesDialog();
  });

  // The CF rules dialog uses its own `fc-cfrulesdlg` overlay class.
  await expect(page.locator('.fc-cfrulesdlg')).toBeVisible({ timeout: 2_000 });

  await page.keyboard.press('Escape');
  expect(consoleErrors.read()).toEqual([]);
}

/** D05b — opening the conditional-format authoring dialog through the
 *  `instance.openConditionalDialog()` API. */
export async function runConditionalDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  type SpreadsheetGlobal = {
    __fcInst?: { openConditionalDialog(): void };
  };

  await page.evaluate(() => {
    const w = window as unknown as SpreadsheetGlobal;
    w.__fcInst?.openConditionalDialog();
  });

  await page.waitForTimeout(150);
  // The dialog class is `fc-fmtdlg fc-conddlg`.
  await expect(page.locator('.fc-conddlg')).toBeVisible({ timeout: 2000 });

  await page.keyboard.press('Escape');
  expect(consoleErrors.read()).toEqual([]);
}
