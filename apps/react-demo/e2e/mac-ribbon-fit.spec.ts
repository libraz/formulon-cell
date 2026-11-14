import { runMacRibbonFitScenario } from '../../../tests/e2e-shared/scenarios/mac-ribbon-fit.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const locale of ['fr', 'de', 'en', 'ja'] as const) {
  test(`Mac: ribbon tabs fit a 1440 wide window as measured (${locale})`, async ({ page }) => {
    await runMacRibbonFitScenario(page, locale);
  });
}
