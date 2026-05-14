import { test } from '@playwright/test';

import {
  runCfRulesDialogScenario,
  runConditionalDialogScenario,
} from '../../../tests/e2e-shared/scenarios/cf-dialog.js';

test('D05 (playground): openCfRulesDialog() surfaces the CF rules overlay', async ({ page }) => {
  await runCfRulesDialogScenario(page);
});

test('D05b (playground): openConditionalDialog() surfaces the authoring overlay', async ({
  page,
}) => {
  await runConditionalDialogScenario(page);
});
