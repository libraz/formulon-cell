import { test } from '@playwright/test';

import { runFormatDialogA11yScenario } from '../../../tests/e2e-shared/scenarios/a11y-dialog.js';

test('A02 (vue-demo): format dialog passes WCAG 2.2 AA', async ({ page }, testInfo) => {
  await runFormatDialogA11yScenario(page, testInfo);
});
