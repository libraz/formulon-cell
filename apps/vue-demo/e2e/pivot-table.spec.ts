import { test } from '@playwright/test';
import { runPivotDateBetweenScenario } from '../../../tests/e2e-shared/scenarios/pivot-date-filter.js';
import {
  runPivotTableFilterDialogScenario,
  runPivotTableRibbonPrimaryDialogScenario,
} from '../../../tests/e2e-shared/scenarios/pivot-table.js';

test('P01 (vue-demo): Label Filters on a row field open the Label Filter dialog', async ({
  page,
}) => {
  await runPivotTableFilterDialogScenario(page);
});

test('P02 (vue-demo): PivotTable ribbon primary click opens Create PivotTable dialog', async ({
  page,
}) => {
  await runPivotTableRibbonPrimaryDialogScenario(page);
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
