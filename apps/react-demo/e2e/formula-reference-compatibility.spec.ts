import { test } from '@playwright/test';
import {
  formulaDataEdgeCases,
  runFormulaDataEdgeNavigation,
  runStructuredReferenceCompletion,
} from '../../../tests/e2e-shared/scenarios/formula-reference-compatibility.js';

for (const scenario of formulaDataEdgeCases) {
  test(`react-demo: ${scenario.platform} ${scenario.surface} ${scenario.style} formula data edge`, async ({
    page,
  }) => {
    await runFormulaDataEdgeNavigation(page, scenario);
  });
}
for (const platform of ['default', 'mac'] as const) {
  for (const surface of ['inline', 'bar'] as const) {
    test(`react-demo: ${platform} ${surface} structured reference completion`, async ({ page }) => {
      await runStructuredReferenceCompletion(page, surface, platform);
    });
  }
}
