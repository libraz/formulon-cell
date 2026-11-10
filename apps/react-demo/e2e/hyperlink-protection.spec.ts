import { test } from '@playwright/test';

import { runProtectedSheetHyperlinkScenario } from '../../../tests/e2e-shared/scenarios/hyperlink-protection.js';

test('a protected sheet takes links on unlocked cells only when Insert hyperlinks is allowed', async ({
  page,
}) => {
  await runProtectedSheetHyperlinkScenario(page);
});
