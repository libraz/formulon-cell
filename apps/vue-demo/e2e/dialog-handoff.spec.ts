import { test } from '@playwright/test';
import {
  runContextDialogReturnScenario,
  runHyperlinkCancelScenario,
  runPasteSpecialCancelScenario,
} from '../../../tests/e2e-shared/scenarios/dialog-handoff.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';

test('context menu dialogs restore grid focus (vue-demo)', async ({ page }) => {
  await runContextDialogReturnScenario(page);
});
test('Paste Special Cancel Enter does not paste (vue-demo)', { tag: SYSTEM_CLIPBOARD }, async ({
  page,
}) => {
  await runPasteSpecialCancelScenario(page);
});
test('Hyperlink Cancel, Remove and Tab keep their own actions (vue-demo)', async ({ page }) => {
  await runHyperlinkCancelScenario(page);
});
