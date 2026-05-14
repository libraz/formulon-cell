import { test } from '@playwright/test';

import { runRibbonInactiveFocusScenario } from '../../../tests/e2e-shared/scenarios/ribbon-tab-focus.js';

test('B5 (react-demo): inactive ribbon panels do not capture Tab focus', async ({ page }) => {
  await runRibbonInactiveFocusScenario(page);
});
