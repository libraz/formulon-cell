import { test } from '@playwright/test';
import { runAxisFormatStructure } from '../../../tests/e2e-shared/scenarios/axis-format-structure.js';
import { runSelectionFormat } from '../../../tests/e2e-shared/scenarios/selection-format.js';

for (const locale of ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th']) {
  for (const platform of ['default', 'mac'] as const) {
    test(`selection formatting ${locale}/${platform}`, async ({ page }) => {
      await runSelectionFormat(page, locale, platform);
    });
    test(`axis formatting after structure ${locale}/${platform}`, async ({ page }) => {
      await runAxisFormatStructure(page, locale, platform);
    });
  }
}
