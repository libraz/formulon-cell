import { test } from '@playwright/test';

import { runMacFindPanelScenario } from '../../../tests/e2e-shared/scenarios/mac-find-panel.js';

test('the Mac find panel steps through matches, wraps, and closes on Escape only with focus', async ({
  page,
}) => {
  await runMacFindPanelScenario(page);
});
