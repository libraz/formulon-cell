import { test } from '@playwright/test';

import { runImeCompositionScenario } from '../../../tests/e2e-shared/scenarios/ime.js';

test('E05 (vue-demo): IME composition sequence does not error', async ({ page }) => {
  await runImeCompositionScenario(page);
});
