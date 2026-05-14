import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

/** V05-V08 — visual snapshots for the per-feature fixtures. Each fixture
 *  applies a deterministic shape (conditional format, sparkline cell, an
 *  active 3×3 selection, or a frozen pane). Baselines are generated on Linux
 *  CI; see playwright.visual.config.ts for the canonical update command. */
const fixtures = [
  { name: 'v05-cf', fixture: 'cf' },
  { name: 'v06-sparkline', fixture: 'sparkline' },
  { name: 'v07-selection', fixture: 'selection' },
  { name: 'v08-frozen', fixture: 'frozen' },
] as const;

for (const { name, fixture } of fixtures) {
  test(`@visual fixture baseline — ${name}`, async ({ page }) => {
    await mountVisualPage(page, `/?fixture=${fixture}`);

    // Let async paint flush before snapping.
    await page.waitForTimeout(400);

    const canvas = page.locator('.fc-host__canvas').first();
    await expect(canvas).toHaveScreenshot(`fixture-${name}.png`, {
      maxDiffPixels: 50,
      animations: 'disabled',
    });
  });
}
