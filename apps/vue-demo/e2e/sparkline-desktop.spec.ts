import { test } from '@playwright/test';

import {
  runSparklineFlatRenderScenario,
  runSparklineHiddenEmptyScenario,
  runSparklineRightToLeftScenario,
  runSparklineStyleStripScenario,
} from '../../../tests/e2e-shared/scenarios/sparkline-desktop.js';

for (const platform of ['default', 'mac'] as const) {
  test(`Edit Data > Hidden & Empty Cells is a submenu on the Mac and a dialog elsewhere (${platform} chrome)`, async ({
    page,
  }) => {
    await runSparklineHiddenEmptyScenario(page, platform);
  });
  test(`Plot Data Right-to-Left mirrors the plot as one undo step (${platform} chrome)`, async ({
    page,
  }) => {
    await runSparklineRightToLeftScenario(page, platform);
  });
}

test('the Sparkline Style strip pages through 29 presets in 80 x 56 chips', async ({ page }) => {
  await runSparklineStyleStripScenario(page);
});

test('flat and single-value sparklines draw as the desktop does', async ({ page }) => {
  await runSparklineFlatRenderScenario(page);
});
