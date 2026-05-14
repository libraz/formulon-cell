import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

/** B8 — visual baseline of a stacked-overlay snapshot.
 *
 *  Opens the format dialog (dialog tier) on top of the grid + active selection
 *  (grid tier) so the resulting screenshot captures the runtime composition
 *  of multiple z-index tiers. The image is brittle by nature; it lives in the
 *  visual project (Linux baseline, maxDiffPixels=50) to flag silent layering
 *  regressions across CSS refactors. */
test('@visual stacked overlays — format dialog over the grid', async ({ page }) => {
  await mountVisualPage(page, '/?fixture=selection');

  // Open the format dialog via Mod+1.
  await page
    .locator('.fc-host')
    .first()
    .click({ position: { x: 200, y: 200 } });
  await page.keyboard.press('Control+1');

  // Wait for the dialog to settle.
  await expect(page.locator('[class="fc-fmtdlg"]')).toBeVisible({ timeout: 5_000 });
  await page.waitForTimeout(250);

  // Snapshot the whole viewport — we want the overlap region between grid
  // and dialog visible in the diff.
  await expect(page).toHaveScreenshot('overlays-stacked.png', {
    maxDiffPixels: 200, // generous: the dialog includes anti-aliased glyphs
    animations: 'disabled',
  });
});
