import { test } from '@playwright/test';
import { runWholeSheetFont } from '../../../tests/e2e-shared/scenarios/whole-sheet-font.js';

for (const platform of ['default', 'mac'] as const) {
  test(`whole-sheet font and Japanese font list: ${platform}`, async ({ page }) => {
    await runWholeSheetFont(page, platform);
  });
}
