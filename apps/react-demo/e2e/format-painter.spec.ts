import { test } from '@playwright/test';

import {
  runCurrencyFormatScenario,
  runFormatPainterScenario,
} from '../../../tests/e2e-shared/scenarios/format-painter.js';

test('F02 (react-demo): apply currency format imperatively without errors', async ({ page }) => {
  await runCurrencyFormatScenario(page);
});

test('F03 (react-demo): Mod+Shift+C/V format painter shortcuts are safe', async ({ page }) => {
  await runFormatPainterScenario(page);
});
