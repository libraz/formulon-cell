import { expect, type Locator, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

async function expectActive(locator: Locator, message: string): Promise<void> {
  await expect
    .poll(async () => locator.evaluate((el) => document.activeElement === el).catch(() => false), {
      message,
    })
    .toBe(true);
}

/** Asserts a modal dialog keeps Tab and Shift+Tab inside it, closes on Escape, and hands focus
 *  back to the grid, as a closed desktop dialog returns to the worksheet. */
async function expectModalFocusContract(page: Page, dialog: Locator) {
  await expect(dialog).toBeVisible();
  const inside = () => dialog.evaluate((el) => el.contains(document.activeElement));
  await expect.poll(inside, { message: 'the dialog should take focus' }).toBe(true);
  for (const key of ['Tab', 'Tab', 'Tab', 'Shift+Tab', 'Shift+Tab', 'Shift+Tab', 'Shift+Tab']) {
    await page.keyboard.press(key);
    expect(await inside(), `${key} should keep focus inside the dialog`).toBe(true);
  }
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expectActive(
    page.locator('.fc-host').first(),
    'the dialog should return focus to the grid',
  );
}

/** B8 — the Review and Automate report and script dialogs follow desktop modal keyboard
 *  semantics: focus moves into the dialog, Tab/Shift+Tab stays inside it, Escape closes, and focus
 *  returns to the grid. */
export async function runDemoModalFocusScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();

  // Spelling opens only on a flagged word.
  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.workbook.batch(() => inst.workbook.sheet(0).setText(0, 0, 'teh'));
    inst.history.clear();
    inst.selection.set('A1');
  });
  await page.locator('[data-ribbon-tab="review"]').click();
  const spelling = page.locator('[data-ribbon-command="spellingReview"]').first();
  await expect(spelling).toBeEnabled();
  await spelling.click();
  await expectModalFocusContract(page, page.getByRole('dialog', { name: /^Spelling/ }));

  await page.locator('[data-ribbon-tab="automate"]').click();
  const script = page.getByRole('button', { name: /^Script$/i }).first();
  await expect(script).toBeEnabled();
  await script.click();
  await expectModalFocusContract(page, page.getByRole('dialog', { name: 'Script', exact: true }));
}

export async function runDemoStatusBarHostStateScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const upload = page.locator('.fc-host__statusbar-upload').first();

  await page.locator('.demo__quick button[aria-label="Save"]').click();
  await expect(upload).toBeVisible();
  await expect(upload).toHaveAttribute('data-upload-status', 'saved');
}
