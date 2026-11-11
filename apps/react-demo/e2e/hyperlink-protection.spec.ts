import { test } from '@playwright/test';

import { runProtectedSheetHyperlinkScenario } from '../../../tests/e2e-shared/scenarios/hyperlink-protection.js';

test('a protected sheet adds and removes links on unlocked cells only with Insert hyperlinks', async ({
  page,
}) => {
  await runProtectedSheetHyperlinkScenario(page);
});
