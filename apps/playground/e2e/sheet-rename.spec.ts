import { test } from '@playwright/test';

import {
  runSheetRenameCancelScenario,
  runSheetRenameScenario,
} from '../../../tests/e2e-shared/scenarios/sheet-rename.js';

test('Sheet rename (playground): F2 → type → Enter renames the active tab', async ({ page }) => {
  await runSheetRenameScenario(page);
});

test('Sheet rename cancel (playground): ESC restores the prior name', async ({ page }) => {
  await runSheetRenameCancelScenario(page);
});
