import {
  runRibbonInactiveFocusScenario,
  runRibbonSplitAltDownScenario,
} from '../../../tests/e2e-shared/scenarios/ribbon-tab-focus.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('B5 (react-demo): inactive ribbon panels do not capture Tab focus', async ({ page }) => {
  await runRibbonInactiveFocusScenario(page);
});

test('B6 (react-demo): Alt+Down opens each split button menu on its first enabled item', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await runRibbonSplitAltDownScenario(page, 'default');
});

test('B6 Mac (react-demo): Alt+Down opens each split button menu on its first enabled item', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await runRibbonSplitAltDownScenario(page, 'mac');
});
