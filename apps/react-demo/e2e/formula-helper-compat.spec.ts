import { test } from '@playwright/test';
import {
  type FormulaHelperSurface,
  formulaHelperVisualCases,
  runFormulaHelperGeometry,
  runFormulaHelperInteraction,
  runFormulaHelperVisual,
  runFormulaKeyboardPointing,
} from '../../../tests/e2e-shared/scenarios/formula-helper-compat.js';

for (const platform of ['default', 'mac'] as const) {
  for (const surface of ['inline', 'bar'] satisfies FormulaHelperSurface[]) {
    test(`react-demo: ${platform} ${surface} function helper interactions`, async ({ page }) => {
      await runFormulaHelperInteraction(page, surface, platform);
    });
    test(`react-demo: ${platform} ${surface} formula keyboard pointing`, async ({ page }) => {
      await runFormulaKeyboardPointing(page, surface, platform);
    });
  }
}
for (const [index, scenario] of formulaHelperVisualCases.entries()) {
  test(`react-demo: helper placement ${index} ${scenario.surface} ${scenario.platform} ${scenario.theme}`, async ({
    page,
  }, info) => {
    const capture = new URL(
      `../../../.scratch/formula-compat/after-react-demo-${info.project.name}-${index}.png`,
      import.meta.url,
    ).pathname;
    await runFormulaHelperVisual(page, scenario, capture);
  });
}

test('react-demo: aligned caret and scaled overlay geometry', async ({ page }) => {
  await runFormulaHelperGeometry(page);
});
