import { runMacShowCommentsPaneScenario } from '../../../tests/e2e-shared/scenarios/mac-comments-pane.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const locale of ['en', 'ja'] as const) {
  test(`Mac: Show Comments toggles the Comments pane (${locale})`, async ({ page }) => {
    await runMacShowCommentsPaneScenario(page, locale);
  });
}
