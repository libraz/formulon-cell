import { test } from '@playwright/test';

import { runSmokeScenario } from '../../../tests/e2e-shared/scenarios/smoke.js';

test('S01/S02: mount cleanly with no console errors and real WASM', async ({ page }) => {
  await runSmokeScenario(page);
});
