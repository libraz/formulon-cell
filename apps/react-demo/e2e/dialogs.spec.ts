import {
  runFindReplaceShortcutScenario,
  runFunctionDialogShortcutScenario,
  runUndoSubsystemSmokeScenario,
} from '../../../tests/e2e-shared/scenarios/dialogs.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('D02 (react-demo): Mod+F opens Find/Replace without errors', async ({ page }) => {
  await runFindReplaceShortcutScenario(page);
});

test('D03 (react-demo): Shift+F3 and fx open Function Arguments and return focus to the sheet', async ({
  page,
}) => {
  await runFunctionDialogShortcutScenario(page);
});

test('D01 smoke (react-demo): Mod+Z reaches the undo subsystem cleanly', async ({ page }) => {
  await runUndoSubsystemSmokeScenario(page);
});
