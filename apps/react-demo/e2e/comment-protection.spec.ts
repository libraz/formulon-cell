import { runProtectedSheetCommentScenario } from '../../../tests/e2e-shared/scenarios/comment-protection.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('a protected sheet allows threaded comments on locked cells and refuses a new note', async ({
  page,
}) => {
  await runProtectedSheetCommentScenario(page);
});
