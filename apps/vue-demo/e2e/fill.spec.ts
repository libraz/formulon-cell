import { test } from '@playwright/test';

import { runFillDownScenario } from '../../../tests/e2e-shared/scenarios/fill.js';

test('Fill-down (vue-demo): Mod+D propagates the anchor value down the selection', async ({
  page,
}) => {
  await runFillDownScenario(page);
});
