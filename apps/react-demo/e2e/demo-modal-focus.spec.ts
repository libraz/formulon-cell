import { test } from '@playwright/test';

import {
  runDemoModalFocusScenario,
  runDemoStatusBarHostStateScenario,
} from '../../../tests/e2e-shared/scenarios/demo-modal-focus.js';

test('B8 (react-demo): wrapper review/script modals trap focus and restore to ribbon', async ({
  page,
}) => {
  await runDemoModalFocusScenario(page);
});

test('B9 (react-demo): demo Save and Script states update the shared status bar', async ({
  page,
}) => {
  await runDemoStatusBarHostStateScenario(page);
});
