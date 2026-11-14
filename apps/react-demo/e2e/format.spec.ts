import {
  runBoldToggleScenario,
  runFormatDialogShortcutScenario,
} from '../../../tests/e2e-shared/scenarios/format.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('F-bold (react-demo): Mod+B toggle leaves the page error-free', async ({ page }) => {
  await runBoldToggleScenario(page);
});

test('F-fmtdlg (react-demo): Mod+1 opens the format dialog', async ({ page }) => {
  await runFormatDialogShortcutScenario(page);
});
