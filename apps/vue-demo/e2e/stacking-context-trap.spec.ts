import { runStackingContextTrapScenario } from '../../../tests/e2e-shared/scenarios/stacking-context-trap.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('C1 (vue-demo): formulon overlay escapes ancestor stacking contexts', async ({ page }) => {
  await runStackingContextTrapScenario(page);
});
