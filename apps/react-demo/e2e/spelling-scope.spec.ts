import { test } from '@playwright/test';

import { runSpellingScopeScenario } from '../../../tests/e2e-shared/scenarios/spelling-scope.js';

test('Spelling checks the whole sheet from one cell and only a selected range otherwise', async ({
  page,
}) => {
  await runSpellingScopeScenario(page);
});
