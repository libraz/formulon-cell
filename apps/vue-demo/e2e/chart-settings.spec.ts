import { test } from '@playwright/test';
import { runChartSettingsScenario } from '../../../tests/e2e-shared/scenarios/chart-settings.js';

test('chart settings and type changes apply with undo and cancel', async ({ page }) => {
  await runChartSettingsScenario(page);
});
