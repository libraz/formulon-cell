import { runWholeSheetFont } from '../../../tests/e2e-shared/scenarios/whole-sheet-font.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`whole-sheet font and Japanese font list: ${platform}`, async ({ page }) => {
    await runWholeSheetFont(page, platform);
  });
}
