import { runValidationDialogScenario } from '../../../tests/e2e-shared/scenarios/validation-dialog.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('D04 (react-demo): format dialog → More tab exposes the validation kind selector', async ({
  page,
}) => {
  await runValidationDialogScenario(page);
});
