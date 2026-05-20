import { test } from '@playwright/test';

import { runExcelChromeBackstageSearchScenario } from '../../../tests/e2e-shared/scenarios/excel-chrome.js';

test('Excel chrome (vue-demo): File backstage, F6 landmarks, and Alt+Q search are wired', async ({
  page,
}) => {
  await runExcelChromeBackstageSearchScenario(page);
});
