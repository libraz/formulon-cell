import { test } from '@playwright/test';

import {
  runBudgetEntryScenario,
  runContactListScenario,
  runMergedSummaryScenario,
  runMonthlySheetsScenario,
  runOrderFillScenario,
  runReportHeaderScenario,
  runSaveOpenContinueScenario,
  runTableOrganizationScenario,
} from '../../../tests/e2e-shared/scenarios/user-journeys.js';

test('user journey: budget entry, currency, formulas, undo, and redo (vue-demo)', async ({
  page,
}) => {
  await runBudgetEntryScenario(page);
});

test('user journey: order formula fill (vue-demo)', async ({ page }) => {
  await runOrderFillScenario(page);
});

test('user journey: contact row insertion and hyperlink editing (vue-demo)', async ({ page }) => {
  await runContactListScenario(page);
});

test('user journey: report header font, alignment, and merge (vue-demo)', async ({ page }) => {
  await runReportHeaderScenario(page);
});

test('user journey: table copy, formats, and structural edits (vue-demo)', async ({ page }) => {
  await runTableOrganizationScenario(page);
});

test('user journey: monthly sheets and cross-sheet copy (vue-demo)', async ({ page }) => {
  await runMonthlySheetsScenario(page);
});

test('user journey: save, open, and continue editing (vue-demo)', async ({ page }) => {
  await runSaveOpenContinueScenario(page);
});

test('user journey: merged selection, AutoSum, and function arguments (vue-demo)', async ({
  page,
}) => {
  await runMergedSummaryScenario(page);
});
