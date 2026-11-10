import { test } from '@playwright/test';

import { runProtectedSheetTableScenario } from '../../../tests/e2e-shared/scenarios/table-protection.js';

test('a protected sheet keeps table banding available and refuses header, insert and Ctrl+T', async ({
  page,
}) => {
  await runProtectedSheetTableScenario(page);
});
