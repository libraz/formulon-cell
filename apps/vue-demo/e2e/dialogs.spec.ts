import { test } from '@playwright/test';

import {
  runFindReplaceShortcutScenario,
  runFunctionDialogShortcutScenario,
  runUndoSubsystemSmokeScenario,
} from '../../../tests/e2e-shared/scenarios/dialogs.js';

test('D02 (vue-demo): Mod+F opens Find/Replace without errors', async ({ page }) => {
  await runFindReplaceShortcutScenario(page);
});

test('D03 (vue-demo): Shift+F3 (fx dialog) is safe', async ({ page }) => {
  await runFunctionDialogShortcutScenario(page);
});

test('D01 smoke (vue-demo): Mod+Z reaches the undo subsystem cleanly', async ({ page }) => {
  await runUndoSubsystemSmokeScenario(page);
});
