import { test } from '@playwright/test';

import {
  runBoldToggleScenario,
  runFormatDialogShortcutScenario,
} from '../../../tests/e2e-shared/scenarios/format.js';

test('F-bold (vue-demo): Mod+B toggle leaves the page error-free', async ({ page }) => {
  await runBoldToggleScenario(page);
});

test('F-fmtdlg (vue-demo): Mod+1 opens the format dialog', async ({ page }) => {
  await runFormatDialogShortcutScenario(page);
});
