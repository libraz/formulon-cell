import { test } from '@playwright/test';
import {
  runFormulaDraftTransactions,
  runR1C1ReferenceTransactions,
} from '../../../tests/e2e-shared/scenarios/formula-edit-transactions.js';

for (const platform of ['default', 'mac'] as const) {
  for (const surface of ['inline', 'bar'] as const) {
    test(`vue-demo: ${platform} ${surface} draft transactions`, async ({ page }) => {
      await runFormulaDraftTransactions(page, surface, platform);
    });
    test(`vue-demo: ${platform} ${surface} R1C1 reference transactions`, async ({ page }) => {
      await runR1C1ReferenceTransactions(page, surface, platform);
    });
  }
}
