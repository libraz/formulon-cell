import { test } from '@playwright/test';

import { runLocaleBootScenario } from '../../../tests/e2e-shared/scenarios/i18n.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';

test('I01 (vue-demo): ?locale=ja boots cleanly without console errors', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runLocaleBootScenario(page);
});
