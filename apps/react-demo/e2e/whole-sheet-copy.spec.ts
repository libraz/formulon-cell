import { test } from '@playwright/test';
import { runWholeSheetCopy } from '../../../tests/e2e-shared/scenarios/whole-sheet-copy.js';

test.skip(({ browserName }) => browserName === 'webkit', 'WebKit denies clipboard reads');

for (const platform of ['default', 'mac'] as const) {
  test(`whole-sheet copy preserves positions and undo: ${platform}`, async ({ page }) => {
    await runWholeSheetCopy(page, platform);
  });
}
