import { test } from '@playwright/test';

import {
  runEditBasicScenario,
  runFormulaScenario,
} from '../../../tests/e2e-shared/scenarios/edit.js';

test('E01: typing a value commits on Enter', async ({ page }) => {
  await runEditBasicScenario(page);
});

test('E02: =SUM(A1:A3) evaluates through real WASM', async ({ page }) => {
  await runFormulaScenario(page);
});
