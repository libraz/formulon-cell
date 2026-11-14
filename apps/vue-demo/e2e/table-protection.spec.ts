import {
  runProtectedSheetCustomStyleScenario,
  runProtectedSheetTableScenario,
} from '../../../tests/e2e-shared/scenarios/table-protection.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`a protected sheet disables the whole Table Design tab, Insert Table and Ctrl+T but keeps Refresh (${platform} chrome)`, async ({
    page,
  }) => {
    await runProtectedSheetTableScenario(page, platform);
  });
  test(`a custom style used on a protected sheet cannot be modified or deleted (${platform} chrome)`, async ({
    page,
  }) => {
    await runProtectedSheetCustomStyleScenario(page, platform);
  });
}
