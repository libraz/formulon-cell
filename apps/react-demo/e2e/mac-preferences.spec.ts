import {
  runMacPreferencesScenario,
  runMacPreferencesViewTitleScenario,
} from '../../../tests/e2e-shared/scenarios/mac-preferences.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Mac Preferences toolbar: Back, Forward, Show All and the pane heading (react-demo)', async ({
  page,
}) => {
  await runMacPreferencesScenario(page);
});

for (const locale of ['en', 'ja', 'de', 'fr', 'zh'] as const) {
  test(`Mac Preferences: the View pane title as drawn (${locale})`, async ({ page }) => {
    await runMacPreferencesViewTitleScenario(page, locale);
  });
}
