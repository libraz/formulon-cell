import { test } from '@playwright/test';

import { runOverlayStackingScenario } from '../../../tests/e2e-shared/scenarios/overlay-stacking.js';

test('B2 (vue-demo): overlays stack on the documented z-index tiers', async ({ page }) => {
  await runOverlayStackingScenario(page);
});
