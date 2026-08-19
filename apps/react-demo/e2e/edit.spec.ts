import { test } from '@playwright/test';

import {
  runEditBasicScenario,
  runFormulaScenario,
  runRecalcAfterEditScenario,
} from '../../../tests/e2e-shared/scenarios/edit.js';

test('E01 (react-demo): typing a value commits on Enter', async ({ page }) => {
  await runEditBasicScenario(page);
});

test('E02 (react-demo): =SUM(A1:A3) evaluates through real WASM', async ({ page }) => {
  await runFormulaScenario(page);
});

test('E06 (react-demo): editing a precedent recomputes its dependents', async ({ page }) => {
  await runRecalcAfterEditScenario(page);
});
