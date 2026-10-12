import { test } from '@playwright/test';
import { runLocaleDialogBehaviorScenario } from '../../../tests/e2e-shared/scenarios/locale-dialog-behavior.js';
import { runLocaleFormulaScenario } from '../../../tests/e2e-shared/scenarios/locale-formulas.js';
import { runLocaleInputScenario } from '../../../tests/e2e-shared/scenarios/locale-input.js';
import { runLocaleShortcutsScenario } from '../../../tests/e2e-shared/scenarios/locale-shortcuts.js';
import { runSevenLocalesScenario } from '../../../tests/e2e-shared/scenarios/seven-locales.js';

test('I02 (react-demo): switching the seven locale profiles updates core and engine', async ({
  page,
}) => {
  await runSevenLocalesScenario(page);
});

test('I03 (react-demo): Mac ribbon and calculation follow all seven locales', async ({ page }) => {
  await runSevenLocalesScenario(page, 'mac');
});

test('I04 (react-demo): localized editing and clipboard preserve numbers and booleans', async ({
  page,
}) => {
  await runLocaleInputScenario(page);
});

for (const platform of ['default', 'mac'] as const) {
  test(`I07 (react-demo): locale dialogs and watch values follow seven locales on ${platform}`, async ({
    page,
  }) => {
    await runLocaleDialogBehaviorScenario(page, platform);
  });
  test(`I06 (react-demo): locale shortcuts and selector focus on ${platform}`, async ({ page }) => {
    await runLocaleShortcutsScenario(page, platform);
  });
  test(`I05 (react-demo): localized formulas round-trip on ${platform}`, async ({ page }) => {
    await runLocaleFormulaScenario(page, platform);
  });
}
