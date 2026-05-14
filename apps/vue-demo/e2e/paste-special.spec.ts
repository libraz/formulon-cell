import { test } from '@playwright/test';

import { runPasteSpecialDialogScenario } from '../../../tests/e2e-shared/scenarios/paste-special.js';

test('C04 (vue-demo): Ctrl+Alt+V opens Paste Special without errors', async ({ page }) => {
  await runPasteSpecialDialogScenario(page);
});
