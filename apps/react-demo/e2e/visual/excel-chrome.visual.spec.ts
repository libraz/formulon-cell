import { test } from '@playwright/test';

import { runExcelChromeVisualScenario } from '../../../../tests/e2e-shared/scenarios/excel-chrome-visual.js';

test.skip(({ browserName }) => browserName !== 'chromium', 'Visual baselines are Chromium-only.');

test('Excel chrome visual baseline (react-demo)', async ({ page }) => {
  await runExcelChromeVisualScenario(page);
});
