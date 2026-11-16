import { runWholeSheetCopy } from '../../../tests/e2e-shared/scenarios/whole-sheet-copy.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`whole-sheet copy preserves positions and undo: ${platform}`, {
    tag: SYSTEM_CLIPBOARD,
  }, async ({ page }) => {
    await runWholeSheetCopy(page, platform);
  });
}
