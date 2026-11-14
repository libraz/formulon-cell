import { runMacNotesMenuScenario } from '../../../tests/e2e-shared/scenarios/notes-menu.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('the Mac Notes menu offers Edit Note on a note cell and drops it while the note is locked', async ({
  page,
}) => {
  await runMacNotesMenuScenario(page);
});
