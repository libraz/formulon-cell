import { test } from '@playwright/test';
import {
  runFreshBookDefaults,
  runStartupFont,
} from '../../../tests/e2e-shared/scenarios/startup-font.js';

for (const locale of ['ja', 'en'] as const) {
  for (const platform of ['default', 'mac'] as const) {
    test(`startup font and readable sample: ${locale}/${platform}`, async ({ page }) => {
      await runStartupFont(page, locale, platform);
    });

    test(`fresh book default column width and row height: ${locale}/${platform}`, async ({
      page,
    }) => {
      await runFreshBookDefaults(page, locale, platform);
    });
  }
}
