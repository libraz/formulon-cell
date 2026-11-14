import { runPasteSpecialDialogScenario } from '../../../tests/e2e-shared/scenarios/paste-special.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('C04 (vue-demo): Ctrl+Alt+V opens Paste Special without errors', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runPasteSpecialDialogScenario(page);
});
