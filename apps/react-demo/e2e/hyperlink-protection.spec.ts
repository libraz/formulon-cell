import { test } from '@playwright/test';

import { runProtectedSheetHyperlinkScenario } from '../../../tests/e2e-shared/scenarios/hyperlink-protection.js';

test('a protected sheet adds links to unlocked cells only with Insert hyperlinks and removes them from unlocked cells', async ({
  page,
}) => {
  await runProtectedSheetHyperlinkScenario(page);
});
