import { test } from '@playwright/test';

import {
  runExcelDataBarAppearanceScenario,
  runExcelDataBarBoundsScenario,
  runExcelIconComparisonScenario,
} from '../../../tests/e2e-shared/scenarios/excel-conditional-parity.js';

for (const locale of ['en', 'ja'] as const) {
  test(`Excel data bar appearance survives edit and XLSX reload (vue-demo, ${locale})`, async ({
    page,
  }) => {
    await runExcelDataBarAppearanceScenario(page, locale);
  });
  test(`Excel data bar bounds survive edit and XLSX reload (vue-demo, ${locale})`, async ({
    page,
  }) => {
    await runExcelDataBarBoundsScenario(page, locale);
  });
  test(`Excel icon comparisons survive edit and XLSX reload (vue-demo, ${locale})`, async ({
    page,
  }) => {
    await runExcelIconComparisonScenario(page, locale);
  });
}
