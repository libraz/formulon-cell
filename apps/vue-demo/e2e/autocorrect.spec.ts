import { test } from '@playwright/test';

import { runAutoCorrectScriptBoundaryScenario } from '../../../tests/e2e-shared/scenarios/autocorrect.js';

test('AutoCorrect replaces Han and kana entries mid-text and Latin entries as whole words', async ({
  page,
}) => {
  await runAutoCorrectScriptBoundaryScenario(page);
});
