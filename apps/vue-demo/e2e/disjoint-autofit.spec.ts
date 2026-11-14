import { runDisjointAutofit } from '../../../tests/e2e-shared/scenarios/disjoint-autofit.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('disjoint header AutoFit and history follow all seven locales', async ({ page }) => {
  test.setTimeout(120_000);
  await runDisjointAutofit(page);
});
