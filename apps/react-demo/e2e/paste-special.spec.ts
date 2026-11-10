import { test } from '@playwright/test';

import { runPasteSpecialDialogScenario } from '../../../tests/e2e-shared/scenarios/paste-special.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';

test('C04 (react-demo): Ctrl+Alt+V opens Paste Special without errors', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runPasteSpecialDialogScenario(page);
});
