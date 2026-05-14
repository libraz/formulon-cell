import { test } from '@playwright/test';

import {
  runLocaleBootScenario,
  runThemeBootScenario,
} from '../../../tests/e2e-shared/scenarios/i18n.js';

test('I01 (playground): ?locale=ja boots cleanly without console errors', async ({ page }) => {
  await runLocaleBootScenario(page);
});

test('I02 (playground): ?theme=dark applies the ink theme to the host', async ({ page }) => {
  await runThemeBootScenario(page);
});
