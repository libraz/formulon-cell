import {
  runSheetTabCopyScenario,
  runSheetTabSwitchScenario,
  runSheetTabsScenario,
} from '../../../tests/e2e-shared/scenarios/sheet-tabs.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('T01 (vue-demo): the + button adds a new sheet and makes it active', async ({ page }) => {
  await runSheetTabsScenario(page);
});

test('T02 (vue-demo): clicking a tab switches the active sheet', async ({ page }) => {
  await runSheetTabSwitchScenario(page);
});

test('T03 (vue-demo): Move or Copy creates a named copy that undo removes and redo restores', async ({
  page,
}) => {
  await runSheetTabCopyScenario(page);
});
