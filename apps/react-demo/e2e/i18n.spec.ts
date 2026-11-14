import { runLocaleBootScenario } from '../../../tests/e2e-shared/scenarios/i18n.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('I01 (react-demo): ?locale=ja boots cleanly without console errors', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runLocaleBootScenario(page);
});
