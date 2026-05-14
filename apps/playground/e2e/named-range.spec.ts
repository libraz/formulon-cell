import { test } from '@playwright/test';

import { runNamedRangeUndoScenario } from '../../../tests/e2e-shared/scenarios/named-range.js';

test('U02 (playground): named range add + undo via imperative __fcInst', async ({ page }) => {
  await runNamedRangeUndoScenario(page);
});
