import { test } from '@playwright/test';
import { runTypedUrlAutoformat } from '../../../tests/e2e-shared/scenarios/typed-url-autoformat.js';

for (const platform of ['default', 'mac'] as const) {
  test(`typed URL formatting and conversion-only undo: ${platform}`, async ({
    page,
    browserName,
  }, testInfo) => {
    await runTypedUrlAutoformat(
      page,
      platform,
      browserName === 'chromium' && platform === 'mac'
        ? async (name) => {
            await page.screenshot({ path: testInfo.outputPath(`${name}.png`) });
          }
        : undefined,
    );
  });
}
