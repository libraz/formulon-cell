import { expect, type Locator, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

async function expectActive(locator: Locator, message: string): Promise<void> {
  await expect
    .poll(async () => locator.evaluate((el) => document.activeElement === el).catch(() => false), {
      message,
    })
    .toBe(true);
}

/** B8 — demo modal dialogs follow desktop spreadsheet keyboard semantics.
 *
 * Review and Automate commands live in the React/Vue wrapper chrome rather
 * than the core spreadsheet. They still need the same Excel-like modal
 * contract as core dialogs: focus moves into the dialog, Tab/Shift+Tab stays
 * inside it, Escape closes, and focus returns to the ribbon command that
 * launched it.
 */
export async function runDemoModalFocusScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();

  await page.locator('[data-ribbon-tab="review"]').click();
  const accessibility = page
    .getByRole('button', { name: /Accessibility|アクセシビリティ/ })
    .first();
  await expect(accessibility).toBeEnabled();
  await accessibility.click();

  const reviewDialog = page
    .locator('.demo__modal[role="dialog"][aria-modal="true"]')
    .filter({ hasText: /Accessibility Check|No issues found/i });
  await expect(reviewDialog).toBeVisible();

  const reviewClose = reviewDialog.getByRole('button', { name: /Close/i });
  const reviewOk = reviewDialog.getByRole('button', { name: /^OK$/i });
  await expectActive(reviewClose, 'review dialog should focus the first command');

  await page.keyboard.press('Shift+Tab');
  await expectActive(reviewOk, 'Shift+Tab should wrap to the last review dialog command');
  await page.keyboard.press('Tab');
  await expectActive(reviewClose, 'Tab should wrap back to the first review dialog command');

  await page.keyboard.press('Escape');
  await expect(reviewDialog).toBeHidden();
  await expectActive(accessibility, 'review dialog should restore focus to its ribbon command');

  await page.locator('[data-ribbon-tab="automate"]').click();
  const script = page.getByRole('button', { name: /^Script$/i }).first();
  await expect(script).toBeEnabled();
  await script.click();

  const scriptDialog = page
    .locator('.demo__modal[role="dialog"][aria-modal="true"]')
    .filter({ hasText: /^Script/ });
  await expect(scriptDialog).toBeVisible();

  const scriptClose = scriptDialog.getByRole('button', { name: /Close/i });
  const scriptRun = scriptDialog.getByRole('button', { name: /^Run$/i });
  await expectActive(scriptClose, 'script dialog should focus the first command');

  await page.keyboard.press('Shift+Tab');
  await expectActive(scriptRun, 'Shift+Tab should wrap to the Run command');
  await page.keyboard.press('Escape');
  await expect(scriptDialog).toBeHidden();
  await expectActive(script, 'script dialog should restore focus to its ribbon command');
}
