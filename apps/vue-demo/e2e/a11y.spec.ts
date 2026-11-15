import {
  runA11yBaselineScenario,
  runA11yEveryTabScenario,
} from '../../../tests/e2e-shared/scenarios/a11y.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('A01 (vue-demo): no WCAG 2.2 AA violations on the mounted app', async ({ page }, testInfo) => {
  await runA11yBaselineScenario(page, testInfo);
});

for (const platform of ['default', 'mac'] as const) {
  test(`A01 (vue-demo): no WCAG 2.2 AA violations with each ribbon tab shown (${platform} chrome)`, async ({
    page,
  }, testInfo) => {
    test.setTimeout(120_000);
    await runA11yEveryTabScenario(page, testInfo, platform);
  });
}
