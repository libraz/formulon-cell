import { runMacFindPanelScenario } from '../../../tests/e2e-shared/scenarios/mac-find-panel.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('the Mac find panel steps through matches, wraps, and closes on Escape only with focus', async ({
  page,
}) => {
  await runMacFindPanelScenario(page);
});
