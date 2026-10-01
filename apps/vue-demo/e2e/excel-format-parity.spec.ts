import { test } from '@playwright/test';

import {
  runExcelFormatCategoryVisibilityScenario,
  runExcelFormatPreviewScenario,
  runExcelGeneralNotationScenario,
} from '../../../tests/e2e-shared/scenarios/excel-format-parity.js';

for (const locale of ['en', 'ja'] as const) {
  test(`Excel General notation previews large and tiny values (vue-demo, ${locale})`, async ({
    page,
  }) => {
    await runExcelGeneralNotationScenario(page, locale);
  });
  test(`Excel Format Cells category visibility (vue-demo, ${locale})`, async ({ page }) => {
    await runExcelFormatCategoryVisibilityScenario(page, locale);
  });
  test(`Excel Format Cells previews the selected value (vue-demo, ${locale})`, async ({ page }) => {
    await runExcelFormatPreviewScenario(page, locale);
  });
}
