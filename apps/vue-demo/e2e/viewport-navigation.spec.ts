import {
  runTabWalkScenario,
  runViewportNavigationScenario,
} from '../../../tests/e2e-shared/scenarios/viewport-navigation.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('viewport selectable and tabBoundary steer keyboard navigation in both modes', async ({
  page,
}) => {
  await runViewportNavigationScenario(page);
});

test('Shift+Tab mirrors Tab, and editable Tab navigation walks only the editable cells', async ({
  page,
}) => {
  await runTabWalkScenario(page);
});
