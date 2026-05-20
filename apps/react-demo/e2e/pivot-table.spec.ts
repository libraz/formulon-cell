import { test } from '@playwright/test';

import {
  runPivotTableFilterDialogScenario,
  runPivotTableRibbonPrimaryDialogScenario,
} from '../../../tests/e2e-shared/scenarios/pivot-table.js';

test('P01 (react-demo): shared PivotTable filter dialog opens from Field Settings', async ({
  page,
}) => {
  await runPivotTableFilterDialogScenario(page);
});

test('P02 (react-demo): PivotTable ribbon primary click opens Create PivotTable dialog', async ({
  page,
}) => {
  await runPivotTableRibbonPrimaryDialogScenario(page);
});
