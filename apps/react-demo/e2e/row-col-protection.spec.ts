import { test } from '@playwright/test';

import { runProtectedSheetRowColumnScenario } from '../../../tests/e2e-shared/scenarios/row-col-protection.js';

test('a protected sheet greys cell shifts and lets whole rows follow their own option', async ({
  page,
}) => {
  await runProtectedSheetRowColumnScenario(page);
});
