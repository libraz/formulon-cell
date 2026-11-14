import { runFormatDialogA11yScenario } from '../../../tests/e2e-shared/scenarios/a11y-dialog.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('A02 (vue-demo): format dialog passes WCAG 2.2 AA', async ({ page }, testInfo) => {
  await runFormatDialogA11yScenario(page, testInfo);
});
