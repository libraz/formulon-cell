import { test } from '@playwright/test';

import {
  runMacEditShortcutScenario,
  runMacPasteSpecialShortcutScenario,
  runMacReferenceShortcutScenario,
} from '../../../tests/e2e-shared/scenarios/mac-operations.js';

test('Mac: Control+U edits, Command+U underlines', async ({ page }) => {
  await runMacEditShortcutScenario(page);
});

test('Mac: Command+T rotates references in both editors', async ({ page }) => {
  await runMacReferenceShortcutScenario(page);
});

test('Mac: Command+Control+V opens Paste Special without pasting', async ({ page }) => {
  await runMacPasteSpecialShortcutScenario(page);
});
