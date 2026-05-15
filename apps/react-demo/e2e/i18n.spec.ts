import { test } from '@playwright/test';

import { runLocaleBootScenario } from '../../../tests/e2e-shared/scenarios/i18n.js';

test('I01 (react-demo): ?locale=ja boots cleanly without console errors', async ({ page }) => {
  await runLocaleBootScenario(page);
});
