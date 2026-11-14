import { runProtectedSheetClearScenario } from '../../../tests/e2e-shared/scenarios/clear-protection.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('a protected sheet greys Clear Contents over any locked cell and Delete changes nothing', async ({
  page,
}) => {
  await runProtectedSheetClearScenario(page);
});
