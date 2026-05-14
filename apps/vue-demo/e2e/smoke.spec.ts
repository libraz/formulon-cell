import { test } from '@playwright/test';

import { runSmokeScenario } from '../../../tests/e2e-shared/scenarios/smoke.js';

test('S01/S02 (vue-demo): mount cleanly with no console errors', async ({ page }) => {
  await runSmokeScenario(page);
});
