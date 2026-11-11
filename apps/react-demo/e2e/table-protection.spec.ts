import { test } from '@playwright/test';

import { runProtectedSheetTableScenario } from '../../../tests/e2e-shared/scenarios/table-protection.js';

for (const platform of ['default', 'mac'] as const) {
  test(`a protected sheet disables the whole Table Design tab, Insert Table and Ctrl+T but keeps Refresh (${platform} chrome)`, async ({
    page,
  }) => {
    await runProtectedSheetTableScenario(page, platform);
  });
}
