import { test } from '@playwright/test';

import {
  runCfRulesDialogScenario,
  runConditionalDialogScenario,
} from '../../../tests/e2e-shared/scenarios/cf-dialog.js';

test('D05 (vue-demo): openCfRulesDialog() surfaces the CF rules overlay', async ({ page }) => {
  await runCfRulesDialogScenario(page);
});

test('D05b (vue-demo): openConditionalDialog() surfaces the authoring overlay', async ({
  page,
}) => {
  await runConditionalDialogScenario(page);
});
