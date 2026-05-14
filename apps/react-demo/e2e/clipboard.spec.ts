import { test } from '@playwright/test';

import {
  runCopyPasteScenario,
  runCutPasteScenario,
} from '../../../tests/e2e-shared/scenarios/clipboard.js';

test('C01 (react-demo): Mod+C/V round-trips a cell value', async ({ page }) => {
  await runCopyPasteScenario(page);
});

test('C02 (react-demo): Mod+X clears the source after paste', async ({ page }) => {
  await runCutPasteScenario(page);
});
