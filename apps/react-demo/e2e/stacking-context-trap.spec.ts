import { test } from '@playwright/test';

import { runStackingContextTrapScenario } from '../../../tests/e2e-shared/scenarios/stacking-context-trap.js';

test('C1 (react-demo): formulon overlay escapes ancestor stacking contexts', async ({ page }) => {
  await runStackingContextTrapScenario(page);
});
