import { runA11yBaselineScenario } from '../../../tests/e2e-shared/scenarios/a11y.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('A01 (vue-demo): no WCAG 2.2 AA violations on the mounted app', async ({ page }, testInfo) => {
  await runA11yBaselineScenario(page, testInfo);
});
