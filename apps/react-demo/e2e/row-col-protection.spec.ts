import { runProtectedSheetRowColumnScenario } from '../../../tests/e2e-shared/scenarios/row-col-protection.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('a protected sheet greys cell shifts and lets whole rows follow their own option', async ({
  page,
}) => {
  await runProtectedSheetRowColumnScenario(page);
});
