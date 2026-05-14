import { test } from '@playwright/test';

import { runDialogMasksGridScenario } from '../../../tests/e2e-shared/scenarios/overlay-hit.js';

test('B1 (playground): dialog overlay masks the grid below', async ({ page }) => {
  await runDialogMasksGridScenario(page);
});
