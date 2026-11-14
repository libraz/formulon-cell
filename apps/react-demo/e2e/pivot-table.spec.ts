import { runPivotDateBetweenScenario } from '../../../tests/e2e-shared/scenarios/pivot-date-filter.js';
import {
  runMacPivotOptionsLayoutScenario,
  runPivotTableFilterDialogScenario,
  runPivotTableOptionsScenario,
  runPivotTableProtectedSheetScenario,
  runPivotTableRibbonPrimaryDialogScenario,
} from '../../../tests/e2e-shared/scenarios/pivot-table.js';
import { test } from '../../../tests/e2e-shared/test.js';

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

test('PivotTable Options edits report and cache options that survive a reopen and an XLSX round trip (react-demo)', async ({
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
