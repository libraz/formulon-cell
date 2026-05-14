import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

/** V01-V04 — paper/ink × en/ja visual baseline of the rendered canvas.
 *
 *  Snapshots are intentionally limited to `.fc-host__canvas` so we don't pick
 *  up the surrounding demo chrome (which can shift across UI tweaks). Baselines
 *  are generated on Linux CI; running locally on macOS or Windows will produce
 *  expected diffs unless you re-baseline with `--update-snapshots` in a docker
 *  container that matches the CI image.
 *
 *  Pixel diff allowance follows the plan's 50px ceiling — that absorbs minor
 *  font-hinting drift across browser versions without masking real regression.
 */
const matrix = [
  { name: 'paper-en', theme: 'light', locale: 'en' },
  { name: 'paper-ja', theme: 'light', locale: 'ja' },
  { name: 'ink-en', theme: 'dark', locale: 'en' },
  { name: 'ink-ja', theme: 'dark', locale: 'ja' },
] as const;

for (const { name, theme, locale } of matrix) {
  test(`@visual grid baseline — ${name}`, async ({ page }) => {
    await mountVisualPage(page, `/?theme=${theme}&locale=${locale}`);

    // Allow async font load + render flush before we snap.
    await page.waitForTimeout(400);

    const canvas = page.locator('.fc-host__canvas').first();
    await expect(canvas).toHaveScreenshot(`grid-${name}.png`, {
      maxDiffPixels: 50,
      animations: 'disabled',
    });
  });
}
