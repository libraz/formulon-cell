import { test } from '@playwright/test';

import { runShowAllNotesScenario } from '../../../tests/e2e-shared/scenarios/show-all-notes.js';

for (const platform of ['default', 'mac'] as const) {
  test(`Show All Notes draws every note without changing the workbook (${platform})`, async ({
    page,
  }) => {
    await runShowAllNotesScenario(page, platform);
  });
}
