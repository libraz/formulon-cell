import { test } from '@playwright/test';

import { runFreezePanesScenario } from '../../../tests/e2e-shared/scenarios/freeze.js';

test('T03 (react-demo): freeze pane toggle round-trip via __fcInst', async ({ page }) => {
  await runFreezePanesScenario(page);
});
