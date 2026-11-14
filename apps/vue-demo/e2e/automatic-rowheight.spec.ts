import { runAutomaticRowheight } from '../../../tests/e2e-shared/scenarios/automatic-rowheight.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`automatic font row height: ${platform}`, async ({ page }) => {
    await runAutomaticRowheight(page, platform);
  });
}
