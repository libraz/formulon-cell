import { test } from '@playwright/test';

import { runDialogMasksGridScenario } from '../../../tests/e2e-shared/scenarios/overlay-hit.js';

test('B1 (vue-demo): dialog overlay masks the grid below', async ({ page }) => {
  await runDialogMasksGridScenario(page);
});
