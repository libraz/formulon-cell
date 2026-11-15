import {
  runSheetBarShiftTabScenario,
  runSheetTabCopyScenario,
  runSheetTabSwitchScenario,
  runSheetTabsScenario,
} from '../../../tests/e2e-shared/scenarios/sheet-tabs.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('T01 (react-demo): the + button adds a new sheet and makes it active', async ({ page }) => {
  await runSheetTabsScenario(page);
});

test('T02 (react-demo): clicking a tab switches the active sheet', async ({ page }) => {
  await runSheetTabSwitchScenario(page);
});

test('T03 (react-demo): Move or Copy creates a named copy that undo removes and redo restores', async ({
  page,
}) => {
  await runSheetTabCopyScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  test(`T04 (react-demo, ${platform}): Shift+Tab moves backwards out of the sheet bar`, async ({
    page,
  }) => {
    await runSheetBarShiftTabScenario(page, platform);
  });
}
