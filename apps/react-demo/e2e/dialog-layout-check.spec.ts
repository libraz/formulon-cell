import { test } from '@playwright/test';

import { runDialogLayoutCheckScenario } from '../../../tests/e2e-shared/scenarios/dialog-layout-check.js';

test('dialog layout check passes self-clipped text and reports overlapping captions', async ({
  page,
}) => {
  await runDialogLayoutCheckScenario(page);
});
