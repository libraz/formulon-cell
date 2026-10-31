import { test } from '@playwright/test';
import { runAutomaticRowheight } from '../../../tests/e2e-shared/scenarios/automatic-rowheight.js';

for (const platform of ['default', 'mac'] as const) {
  test(`automatic font row height: ${platform}`, async ({ page, browserName }, testInfo) => {
    await runAutomaticRowheight(
      page,
      platform,
      platform === 'mac' && browserName === 'chromium'
        ? async () => {
            await page.screenshot({ path: testInfo.outputPath('large-font.png') });
          }
        : undefined,
    );
  });
}
