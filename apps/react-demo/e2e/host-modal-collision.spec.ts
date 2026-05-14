import { test } from '@playwright/test';

import { runHostModalCollisionScenario } from '../../../tests/e2e-shared/scenarios/host-modal-collision.js';

test('B3 (react-demo): formulon overlays stay above a z-9999 host modal', async ({ page }) => {
  await runHostModalCollisionScenario(page);
});
