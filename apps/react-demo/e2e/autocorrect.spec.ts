import { runAutoCorrectScriptBoundaryScenario } from '../../../tests/e2e-shared/scenarios/autocorrect.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('AutoCorrect replaces Han and kana entries mid-text and Latin entries as whole words', async ({
  page,
}) => {
  await runAutoCorrectScriptBoundaryScenario(page);
});
