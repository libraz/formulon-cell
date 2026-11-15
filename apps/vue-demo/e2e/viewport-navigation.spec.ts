import { runViewportNavigationScenario } from '../../../tests/e2e-shared/scenarios/viewport-navigation.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('viewport selectable and tabBoundary steer keyboard navigation in both modes', async ({
  page,
}) => {
  await runViewportNavigationScenario(page);
});
