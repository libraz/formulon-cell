import { test } from '@playwright/test';

import { runProtectedSheetObjectScenario } from '../../../tests/e2e-shared/scenarios/object-protection.js';

test('a protected sheet offers chart and shape insertion only with Edit objects', async ({
  page,
}) => {
  await runProtectedSheetObjectScenario(page);
});
