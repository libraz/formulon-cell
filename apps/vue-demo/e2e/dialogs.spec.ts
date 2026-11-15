import {
  runFindReplaceShortcutScenario,
  runFunctionDialogShortcutScenario,
  runMacFormulaBuilderKeyboardScenario,
  runUndoSubsystemSmokeScenario,
} from '../../../tests/e2e-shared/scenarios/dialogs.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('D02 (vue-demo): Mod+F opens Find/Replace without errors', async ({ page }) => {
  await runFindReplaceShortcutScenario(page);
});

test('D03 (vue-demo): Shift+F3 and fx open Function Arguments and return focus to the sheet', async ({
  page,
}) => {
  await runFunctionDialogShortcutScenario(page);
});

test('D03 Mac (vue-demo): Shift+F3 opens the Formula Builder pane, Tab leaves it open, close returns to the sheet', async ({
  page,
}) => {
  await runMacFormulaBuilderKeyboardScenario(page);
});

test('D01 smoke (vue-demo): Mod+Z reaches the undo subsystem cleanly', async ({ page }) => {
  await runUndoSubsystemSmokeScenario(page);
});
