import { test } from '@playwright/test';

import {
  runFilterShortcutScenario,
  runProtectedSheetFilterAndBackgroundScenario,
  runProtectedSheetFilterButtonScenario,
} from '../../../tests/e2e-shared/scenarios/filter-protection.js';

test('Ctrl+Shift+L toggles AutoFilter over the current region', async ({ page }) => {
  await runFilterShortcutScenario(page);
});

test('a protected sheet disables Background and refuses Ctrl+Shift+L', async ({ page }) => {
  await runProtectedSheetFilterAndBackgroundScenario(page);
});

test('a protected sheet leaves a filter button inert without Use AutoFilter', async ({ page }) => {
  await runProtectedSheetFilterButtonScenario(page);
});
