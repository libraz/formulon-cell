import { test } from '@playwright/test';

import { runFormatDialogFocusTrapScenario } from '../../../tests/e2e-shared/scenarios/format-dialog-focus-trap.js';

test('B7 (vue-demo): format dialog focus stays in the active tab', async ({ page }) => {
  await runFormatDialogFocusTrapScenario(page);
});
