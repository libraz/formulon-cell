import { runLocaleVisualLayout } from '../../../tests/e2e-shared/scenarios/locale-visual-layout.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  for (const width of [1280, 768]) {
    test(`locale visual layout: ${platform} ${width}px`, async ({ page }, info) => {
      test.setTimeout(180_000);
      await runLocaleVisualLayout(page, info, platform, width);
    });
  }
}
