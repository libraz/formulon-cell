import { test } from '@playwright/test';

import { runProtectedSheetClearScenario } from '../../../tests/e2e-shared/scenarios/clear-protection.js';

test('a protected sheet greys Clear Contents over any locked cell and Delete changes nothing', async ({
  page,
}) => {
  await runProtectedSheetClearScenario(page);
});
