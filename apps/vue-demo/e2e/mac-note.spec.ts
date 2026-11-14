import { runMacNewNoteScenario } from '../../../tests/e2e-shared/scenarios/mac-note.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Mac: a new note is a bare 104x76 pt box beside the cell', async ({ page }) => {
  await runMacNewNoteScenario(page);
});
