import { test } from '@playwright/test';
import { runStartupFont } from '../../../tests/e2e-shared/scenarios/startup-font.js';

for (const locale of ['ja', 'en'] as const) {
  for (const platform of ['default', 'mac'] as const) {
    test(`startup font and readable sample: ${locale}/${platform}`, async ({ page }) => {
      await runStartupFont(page, locale, platform);
    });
  }
}
