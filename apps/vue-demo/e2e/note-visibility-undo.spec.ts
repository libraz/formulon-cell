import { runNoteVisibilityUndoScenario } from '../../../tests/e2e-shared/scenarios/note-visibility-undo.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  test(`showing or hiding notes is not an undo step (${platform})`, async ({ page }) => {
    await runNoteVisibilityUndoScenario(page, platform);
  });
}
