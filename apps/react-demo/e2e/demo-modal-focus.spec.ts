import { test } from '@playwright/test';

import {
  runDemoModalFocusScenario,
  runDemoStatusBarHostStateScenario,
} from '../../../tests/e2e-shared/scenarios/demo-modal-focus.js';

test('B8 (react-demo): Review and Automate modals trap focus and return it to the grid', async ({
  page,
}) => {
  await runDemoModalFocusScenario(page);
});

test('B9 (react-demo): demo Save updates the shared status bar', async ({ page }) => {
  await runDemoStatusBarHostStateScenario(page);
});
