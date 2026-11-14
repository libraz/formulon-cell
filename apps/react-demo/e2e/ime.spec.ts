import { runImeCompositionScenario } from '../../../tests/e2e-shared/scenarios/ime.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('E05 (react-demo): IME composition sequence does not error', async ({ page }) => {
  await runImeCompositionScenario(page);
});
