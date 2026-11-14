import { runMacPreferencesScenario } from '../../../tests/e2e-shared/scenarios/mac-preferences.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Mac Preferences toolbar: Back, Forward, Show All and the pane heading (vue-demo)', async ({
  page,
}) => {
  await runMacPreferencesScenario(page);
});
