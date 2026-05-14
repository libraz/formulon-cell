import { test } from '@playwright/test';

import { runUndoRedoScenario } from '../../../tests/e2e-shared/scenarios/undo-redo.js';

test('U01 (react-demo): edit → undo → redo via Mod+Z / Mod+Shift+Z', async ({ page }) => {
  await runUndoRedoScenario(page);
});
