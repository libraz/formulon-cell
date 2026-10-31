import { test } from '@playwright/test';
import { runTypedUrlAutoformat } from '../../../tests/e2e-shared/scenarios/typed-url-autoformat.js';

for (const platform of ['default', 'mac'] as const) {
  test(`typed URL formatting and conversion-only undo: ${platform}`, async ({ page }) => {
    await runTypedUrlAutoformat(page, platform);
  });
}
