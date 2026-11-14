import { runNamedRangeUndoScenario } from '../../../tests/e2e-shared/scenarios/named-range.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('U02 (vue-demo): named range add + undo via imperative __fcInst', async ({ page }) => {
  await runNamedRangeUndoScenario(page);
});
