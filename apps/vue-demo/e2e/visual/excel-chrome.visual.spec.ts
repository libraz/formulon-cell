import { runExcelChromeVisualScenario } from '../../../../tests/e2e-shared/scenarios/excel-chrome-visual.js';
import { SYSTEM_CLIPBOARD } from '../../../../tests/e2e-shared/tags.js';
import { test } from '../../../../tests/e2e-shared/test.js';

test.skip(({ browserName }) => browserName !== 'chromium', 'Visual baselines are Chromium-only.');

test('Excel chrome visual baseline (vue-demo)', { tag: SYSTEM_CLIPBOARD }, async ({ page }) => {
  // One scenario walks every ribbon tab, menu and dialog for ~90 screenshots.
  test.setTimeout(180_000);
  await runExcelChromeVisualScenario(page);
});
