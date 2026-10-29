import { test } from '@playwright/test';
import {
  runCfRulesDialogScenario,
  runConditionalDialogScenario,
} from '../../../tests/e2e-shared/scenarios/cf-dialog.js';
import { runCfRulesKeyboardScenario } from '../../../tests/e2e-shared/scenarios/cf-rules-keyboard.js';

test('D05 (react-demo): openCfRulesDialog() surfaces the CF rules overlay', async ({ page }) => {
  await runCfRulesDialogScenario(page);
});

test('D05b (react-demo): openConditionalDialog() surfaces the authoring overlay', async ({
  page,
}) => {
  await runConditionalDialogScenario(page);
});

test('D05c (react-demo): CF rules controls keep native keyboard ownership', async ({ page }) => {
  await runCfRulesKeyboardScenario(page);
});
