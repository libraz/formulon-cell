import {
  runDemoModalFocusScenario,
  runDemoStatusBarHostStateScenario,
} from '../../../tests/e2e-shared/scenarios/demo-modal-focus.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('B8 (vue-demo): Review and Automate modals trap focus and return it to the grid', async ({
  page,
}) => {
  await runDemoModalFocusScenario(page);
});

test('B9 (vue-demo): demo Save updates the shared status bar', async ({ page }) => {
  await runDemoStatusBarHostStateScenario(page);
});
