import { test } from '@playwright/test';

import {
  runSheetTabSwitchScenario,
  runSheetTabsScenario,
} from '../../../tests/e2e-shared/scenarios/sheet-tabs.js';

test('T01 (playground): the + button adds a new sheet and makes it active', async ({ page }) => {
  await runSheetTabsScenario(page);
});

test('T02 (playground): clicking a tab switches the active sheet', async ({ page }) => {
  await runSheetTabSwitchScenario(page);
});
