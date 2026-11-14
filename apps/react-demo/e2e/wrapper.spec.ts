import { runWrapperReactivityScenario } from '../../../tests/e2e-shared/scenarios/wrapper-reactivity.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('useSelection (react-demo): the Selection card re-renders when the active cell moves', async ({
  page,
}) => {
  await runWrapperReactivityScenario(page);
});
