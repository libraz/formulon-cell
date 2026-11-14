import { runDialogMasksGridScenario } from '../../../tests/e2e-shared/scenarios/overlay-hit.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('B1 (react-demo): dialog overlay masks the grid below', async ({ page }) => {
  await runDialogMasksGridScenario(page);
});
