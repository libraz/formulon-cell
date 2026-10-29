import { test } from '@playwright/test';
import {
  type FormulaHelperSurface,
  formulaHelperVisualCases,
  runFormulaHelperGeometry,
  runFormulaHelperInteraction,
  runFormulaHelperVisual,
} from '../../../tests/e2e-shared/scenarios/formula-helper-compat.js';

for (const platform of ['default', 'mac'] as const) {
  for (const surface of ['inline', 'bar'] satisfies FormulaHelperSurface[]) {
    test(`vue-demo: ${platform} ${surface} function helper interactions`, async ({ page }) => {
      await runFormulaHelperInteraction(page, surface, platform);
    });
  }
}
for (const [index, scenario] of formulaHelperVisualCases.entries()) {
  test(`vue-demo: helper placement ${index} ${scenario.surface} ${scenario.platform} ${scenario.theme}`, async ({
    page,
  }, info) => {
    const capture = new URL(
      `../../../.scratch/formula-compat/after-vue-demo-${info.project.name}-${index}.png`,
      import.meta.url,
    ).pathname;
    await runFormulaHelperVisual(page, scenario, capture);
  });
}

test('vue-demo: aligned caret and scaled overlay geometry', async ({ page }) => {
  await runFormulaHelperGeometry(page);
});
