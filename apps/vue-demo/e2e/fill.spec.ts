import { test } from '@playwright/test';

import {
  runFillDownScenario,
  runRibbonFillDownScenario,
} from '../../../tests/e2e-shared/scenarios/fill.js';

test('Fill-down (vue-demo): Mod+D propagates the anchor value down the selection', async ({
  page,
}) => {
  await runFillDownScenario(page);
});

// Phase 2: ribbon Fill > Down used to silent-fail in Vue because
// `createDynamicDropdowns` was never wired. mountToolbar now auto-attaches a
// click delegator when the host passes `dynamicDropdowns: true`.
test('Ribbon Fill (vue-demo): clicking the Fill > Down menu fills the selected range', async ({
  page,
}) => {
  await runRibbonFillDownScenario(page);
});
