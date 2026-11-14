import { runHostModalCollisionScenario } from '../../../tests/e2e-shared/scenarios/host-modal-collision.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('B3 (vue-demo): formulon overlays stay above a z-9999 host modal', async ({ page }) => {
  await runHostModalCollisionScenario(page);
});
