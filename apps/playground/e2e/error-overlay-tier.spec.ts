import { test } from '@playwright/test';

import { runErrorOverlayTierScenario } from '../../../tests/e2e-shared/scenarios/error-overlay-tier.js';

test('B9 (playground): .fc-errmenu lives on the top z-tier', async ({ page }) => {
  await runErrorOverlayTierScenario(page);
});
