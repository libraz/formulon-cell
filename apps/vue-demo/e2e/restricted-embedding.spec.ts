import { test } from '@playwright/test';
import { runRestrictedEmbeddingScenario } from '../../../tests/e2e-shared/scenarios/restricted-embedding.js';

test('restricted embedding in a host modal (Vue)', async ({ page }) => {
  await runRestrictedEmbeddingScenario(page);
});
