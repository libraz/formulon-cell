import { test } from '@playwright/test';

import {
  runCurrencyFormatScenario,
  runFormatPainterScenario,
} from '../../../tests/e2e-shared/scenarios/format-painter.js';
import { runStickyFormatPainterScenario } from '../../../tests/e2e-shared/scenarios/format-painter-sticky.js';

test('F02 (vue-demo): apply currency format imperatively without errors', async ({ page }) => {
  await runCurrencyFormatScenario(page);
});

test('F03 (vue-demo): Mod+Shift+C/V format painter shortcuts are safe', async ({ page }) => {
  await runFormatPainterScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  test(`Format Painter button paints once on click and keeps painting after a double-click (${platform} chrome)`, async ({
    page,
  }) => {
    await runStickyFormatPainterScenario(page, platform);
  });
}
