import { runRibbonInactiveFocusScenario } from '../../../tests/e2e-shared/scenarios/ribbon-tab-focus.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('B5 (react-demo): inactive ribbon panels do not capture Tab focus', async ({ page }) => {
  await runRibbonInactiveFocusScenario(page);
});
