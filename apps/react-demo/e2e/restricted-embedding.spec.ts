import { runRestrictedEmbeddingScenario } from '../../../tests/e2e-shared/scenarios/restricted-embedding.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('restricted embedding in a host modal (React)', async ({ page }) => {
  await runRestrictedEmbeddingScenario(page);
});
