import { runHeaderClickActiveCellScenario } from '../../../tests/e2e-shared/scenarios/header-click-active-cell.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('a header click puts the active cell on the first visible cell without scrolling', async ({
  page,
}) => {
  await runHeaderClickActiveCellScenario(page);
});
