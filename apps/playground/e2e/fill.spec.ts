import { test } from '@playwright/test';

import {
  runFillDownScenario,
  runRibbonFillDownScenario,
} from '../../../tests/e2e-shared/scenarios/fill.js';

test('Fill-down (playground): Mod+D propagates the anchor value down the selection', async ({
  page,
}) => {
  await runFillDownScenario(page);
});

test('Ribbon Fill (playground): clicking the Fill > Down menu fills the selected range', async ({
  page,
}) => {
  await runRibbonFillDownScenario(page);
});
