import { test } from '@playwright/test';
import { runPivotDateBetweenScenario } from '../../../tests/e2e-shared/scenarios/pivot-date-filter.js';
import {
  runPivotTableFilterDialogScenario,
  runPivotTableRibbonPrimaryDialogScenario,
} from '../../../tests/e2e-shared/scenarios/pivot-table.js';

test('P01 (vue-demo): shared PivotTable filter dialog opens from Field Settings', async ({
  page,
}) => {
  await runPivotTableFilterDialogScenario(page);
});

test('P02 (vue-demo): PivotTable ribbon primary click opens Create PivotTable dialog', async ({
  page,
}) => {
  await runPivotTableRibbonPrimaryDialogScenario(page);
});

for (const locale of ['en', 'de', 'ja'] as const) {
  test(`PivotTable Date is between reads ${locale} dates and keeps only the rows inside them`, async ({
    page,
  }) => {
    await runPivotDateBetweenScenario(page, locale);
  });
}
