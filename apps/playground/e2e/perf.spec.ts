import { test } from '@playwright/test';

import { runWideSelectionPerfScenario } from '../../../tests/e2e-shared/scenarios/perf.js';

test('N03 (playground): a wide selection update is fast (< 2s for 10k cells)', async ({ page }) => {
  await runWideSelectionPerfScenario(page);
});
