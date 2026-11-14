import { runMacErrorCheckingPaneScenario } from '../../../tests/e2e-shared/scenarios/mac-error-checking-pane.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const locale of ['en', 'ja'] as const) {
  test(`Mac Preferences: Error Checking pane rules, master switch and reset (${locale})`, async ({
    page,
  }) => {
    await runMacErrorCheckingPaneScenario(page, locale);
  });
}
