import {
  runPickListContentScenario,
  runProtectedSheetMenuScenario,
} from '../../../tests/e2e-shared/scenarios/menu-protection.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('a protected sheet greys the cell menu by option and by locked selection', async ({
  page,
}) => {
  await runProtectedSheetMenuScenario(page);
});

test('Pick From Drop-down List offers the column block text without the active cell', async ({
  page,
}) => {
  await runPickListContentScenario(page);
});
