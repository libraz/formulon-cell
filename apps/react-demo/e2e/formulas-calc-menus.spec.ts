import { runCalcOptionsMenusScenario } from '../../../tests/e2e-shared/scenarios/formulas-calc-menus.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`Formulas: Circular References, Format Stale Values and Compatibility Version menus (${platform} chrome)`, async ({
    page,
  }) => {
    await runCalcOptionsMenusScenario(page, platform);
  });
}
