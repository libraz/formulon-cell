import { runErrorOverlayTierScenario } from '../../../tests/e2e-shared/scenarios/error-overlay-tier.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('B9 (vue-demo): .fc-errmenu lives on the top z-tier', async ({ page }) => {
  await runErrorOverlayTierScenario(page);
});
