import { test } from '@playwright/test';

import { runDefaultNoteMenuScenario } from '../../../tests/e2e-shared/scenarios/default-note-menu.js';

test('default profile offers New Note on a plain cell and the note items on a noted one', async ({
  page,
}) => {
  await runDefaultNoteMenuScenario(page);
});
