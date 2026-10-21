import { test } from '@playwright/test';
import { runPlatformGeometry } from '../../../tests/e2e-shared/scenarios/platform-geometry.js';

for (const locale of ['ja', 'en'] as const) {
  for (const platform of ['default', 'mac'] as const) {
    test(`blank workbook screen geometry: ${locale}/${platform}`, async ({ page }) => {
      await runPlatformGeometry(page, locale, platform);
    });
  }
}
