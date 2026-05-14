import { test } from '@playwright/test';

import {
  runDivByZeroScenario,
  runUnknownFunctionScenario,
} from '../../../tests/e2e-shared/scenarios/error.js';

test('N01 (playground): =1/0 yields #DIV/0! without throwing', async ({ page }) => {
  await runDivByZeroScenario(page);
});

test('N02 (playground): =NOTAFN() yields #NAME? without throwing', async ({ page }) => {
  await runUnknownFunctionScenario(page);
});
