import { test } from '@playwright/test';
import { runPivotDateBetweenScenario } from '../../../tests/e2e-shared/scenarios/pivot-date-filter.js';
import {
  runPivotTableFilterDialogScenario,
  runPivotTableProtectedSheetScenario,
  runPivotTableRibbonPrimaryDialogScenario,
} from '../../../tests/e2e-shared/scenarios/pivot-table.js';

test('P01 (react-demo): Label Filters on a row field open the Label Filter dialog', async ({
  page,
}) => {
  await runPivotTableFilterDialogScenario(page);
});

test('P02 (react-demo): PivotTable ribbon primary click opens Create PivotTable dialog', async ({
  page,
}) => {
  await runPivotTableRibbonPrimaryDialogScenario(page);
});

test('P03 (react-demo): a protected sheet makes Insert PivotTable unavailable until unprotected', async ({
  page,
}) => {
  await runPivotTableProtectedSheetScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  for (const locale of ['en', 'de', 'ja'] as const) {
    test(`PivotTable Date Filters read ${locale} dates for Between and apply a period at once (${platform} chrome)`, async ({
      page,
    }) => {
      await runPivotDateBetweenScenario(page, locale, platform);
    });
  }
}
