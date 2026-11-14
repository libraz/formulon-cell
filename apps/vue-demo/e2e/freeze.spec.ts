import { runFreezePanesScenario } from '../../../tests/e2e-shared/scenarios/freeze.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('T03 (vue-demo): freeze pane toggle round-trip via __fcInst', async ({ page }) => {
  await runFreezePanesScenario(page);
});
