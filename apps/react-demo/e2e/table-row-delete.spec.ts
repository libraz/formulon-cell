import { runTableRowDeleteScenario } from '../../../tests/e2e-shared/scenarios/table-row-delete.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('deleting table rows keeps the table in step and undo restores it exactly', async ({
  page,
}) => {
  await runTableRowDeleteScenario(page);
});
