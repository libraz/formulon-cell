import {
  runLocaleBootScenario,
  runLocalePasteSpecialScenario,
} from '../../../tests/e2e-shared/scenarios/i18n.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('I01 (react-demo): ?locale=ja boots cleanly without console errors', async ({ page }) => {
  await runLocaleBootScenario(page);
});

test('I01 (react-demo): ?locale=ja localizes Paste Special after a copy', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runLocalePasteSpecialScenario(page);
});
