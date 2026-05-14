import { test } from '@playwright/test';

import { runWrapperReactivityScenario } from '../../../tests/e2e-shared/scenarios/wrapper-reactivity.js';

test('useSelection (vue-demo): the Selection card re-renders when the active cell moves', async ({
  page,
}) => {
  await runWrapperReactivityScenario(page);
});
