import { runAutomaticRowheight } from '../../../tests/e2e-shared/scenarios/automatic-rowheight.js';
import { test } from '../../../tests/e2e-shared/test.js';

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
