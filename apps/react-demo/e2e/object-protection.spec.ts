import {
  runProtectedSheetObjectScenario,
  runProtectedSheetSlicerScenario,
} from '../../../tests/e2e-shared/scenarios/object-protection.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('a protected sheet offers chart and shape insertion only with Edit objects', async ({
  page,
}) => {
  await runProtectedSheetObjectScenario(page);
});

test('a protected sheet decides slicer clicks by Edit objects and Use AutoFilter', async ({
  page,
}) => {
  await runProtectedSheetSlicerScenario(page);
});
