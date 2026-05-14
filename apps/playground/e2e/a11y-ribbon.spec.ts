import { test } from '@playwright/test';

import { runRibbonA11yScenario } from '../../../tests/e2e-shared/scenarios/a11y-ribbon.js';

test('A03 (playground): ribbon/toolbar passes WCAG 2.2 AA', async ({ page }, testInfo) => {
  await runRibbonA11yScenario(page, testInfo);
});
