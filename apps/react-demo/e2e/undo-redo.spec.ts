import { runUndoRedoScenario } from '../../../tests/e2e-shared/scenarios/undo-redo.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('U01 (react-demo): edit → undo → redo via Mod+Z / Mod+Shift+Z', async ({ page }) => {
  await runUndoRedoScenario(page);
});
