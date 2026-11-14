import { runPivotDateBetweenScenario } from '../../../tests/e2e-shared/scenarios/pivot-date-filter.js';
import {
  runMacPivotOptionsLayoutScenario,
  runPivotTableFilterDialogScenario,
  runPivotTableOptionsScenario,
  runPivotTableRibbonPrimaryDialogScenario,
} from '../../../tests/e2e-shared/scenarios/pivot-table.js';
import { test } from '../../../tests/e2e-shared/test.js';

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

test('PivotTable Options edits report and cache options that survive a reopen and an XLSX round trip (vue-demo)', async ({
  page,
}) => {
  await runPivotTableOptionsScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  for (const locale of ['en', 'de', 'ja'] as const) {
    test(`PivotTable Date Filters read ${locale} dates for Between and All Dates in the Period (${platform} chrome)`, async ({
      page,
    }) => {
      await runPivotDateBetweenScenario(page, locale, platform);
    });
  }
}

for (const locale of ['en', 'ja'] as const) {
  test(`Mac PivotTable Options greys Show items with no data and has no table sort (${locale})`, async ({
    page,
  }) => {
    await runMacPivotOptionsLayoutScenario(page, locale);
  });
}
