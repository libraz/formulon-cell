import {
  runSheetRenameCancelScenario,
  runSheetRenameHistoryScenario,
  runSheetRenameScenario,
} from '../../../tests/e2e-shared/scenarios/sheet-rename.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Sheet rename (react-demo): F2 → type → Enter renames the active tab', async ({ page }) => {
  await runSheetRenameScenario(page);
});

test('Sheet rename cancel (react-demo): ESC restores the prior name', async ({ page }) => {
  await runSheetRenameCancelScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  test(`Sheet rename History (react-demo): reserved off the Mac, accepted on it (${platform} chrome)`, async ({
    page,
  }) => {
    await runSheetRenameHistoryScenario(page, platform);
  });
}
