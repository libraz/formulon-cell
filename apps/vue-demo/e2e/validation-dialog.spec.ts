import { test } from '@playwright/test';

import { runValidationDialogScenario } from '../../../tests/e2e-shared/scenarios/validation-dialog.js';

test('D04 (vue-demo): format dialog → More tab exposes the validation kind selector', async ({
  page,
}) => {
  await runValidationDialogScenario(page);
});
