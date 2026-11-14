import { runCellShiftModalScenario } from '../../../tests/e2e-shared/scenarios/cell-shift-modal.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const kind of ['insert', 'delete'] as const) {
  test(`cell ${kind} dialog keeps focus and cancel semantics (react-demo)`, async ({ page }) => {
    await runCellShiftModalScenario(page, kind);
  });
}
