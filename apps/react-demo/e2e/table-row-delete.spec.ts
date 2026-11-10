import { test } from '@playwright/test';

import { runTableRowDeleteScenario } from '../../../tests/e2e-shared/scenarios/table-row-delete.js';

test('deleting table rows keeps the table in step and undo restores it exactly', async ({
  page,
}) => {
  await runTableRowDeleteScenario(page);
});
