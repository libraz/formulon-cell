import { test } from '@playwright/test';
import { runAutomaticRowheight } from '../../../tests/e2e-shared/scenarios/automatic-rowheight.js';

for (const platform of ['default', 'mac'] as const) {
  test(`automatic font row height: ${platform}`, async ({ page }) => {
    await runAutomaticRowheight(page, platform);
  });
}
