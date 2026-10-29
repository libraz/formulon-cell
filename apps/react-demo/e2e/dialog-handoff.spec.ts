import { test } from '@playwright/test';
import {
  runContextDialogReturnScenario,
  runHyperlinkCancelScenario,
  runPasteSpecialCancelScenario,
} from '../../../tests/e2e-shared/scenarios/dialog-handoff.js';

test('context menu dialogs restore grid focus (react-demo)', async ({ page }) => {
  await runContextDialogReturnScenario(page);
});
test('Paste Special Cancel Enter does not paste (react-demo)', async ({ page }) => {
  await runPasteSpecialCancelScenario(page);
});
test('Hyperlink Cancel, Remove and Tab keep their own actions (react-demo)', async ({ page }) => {
  await runHyperlinkCancelScenario(page);
});
