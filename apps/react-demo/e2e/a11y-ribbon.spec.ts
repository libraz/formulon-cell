import { runRibbonA11yScenario } from '../../../tests/e2e-shared/scenarios/a11y-ribbon.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('A03 (react-demo): ribbon/toolbar passes WCAG 2.2 AA', async ({ page }, testInfo) => {
  await runRibbonA11yScenario(page, testInfo);
});
