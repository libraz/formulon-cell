import { test } from '@playwright/test';
import {
  runNativeImageSaveScenario,
  runNativeR1C1Scenario,
  runNativeThreadThemeScenario,
} from '../../../tests/e2e-shared/scenarios/upstream-013.js';

test('native R1C1 display and editing', async ({ page }) => {
  await runNativeR1C1Scenario(page);
});

test('native threaded comments and workbook theme from Mac ribbon', async ({ page }) => {
  await runNativeThreadThemeScenario(page);
});

test('native image insertion and save/open', async ({ page }) => {
  await runNativeImageSaveScenario(page);
});
