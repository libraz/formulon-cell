import {
  runContextDialogReturnScenario,
  runHyperlinkCancelScenario,
  runPasteSpecialCancelScenario,
} from '../../../tests/e2e-shared/scenarios/dialog-handoff.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('context menu dialogs restore grid focus (react-demo)', async ({ page }) => {
  await runContextDialogReturnScenario(page);
});
test('Paste Special Cancel Enter does not paste (react-demo)', { tag: SYSTEM_CLIPBOARD }, async ({
  page,
}) => {
  await runPasteSpecialCancelScenario(page);
});
test('Hyperlink Cancel, Remove and Tab keep their own actions (react-demo)', async ({ page }) => {
  await runHyperlinkCancelScenario(page);
});
