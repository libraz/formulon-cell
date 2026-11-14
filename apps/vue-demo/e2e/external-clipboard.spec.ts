import { runExternalHtmlPasteScenario } from '../../../tests/e2e-shared/scenarios/external-clipboard.js';
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('C03 (vue-demo): external HTML payload pastes via system clipboard', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page, context, browserName }) => {
  test.skip(browserName === 'webkit', 'WebKit denies clipboard.write without a user gesture');
  await runExternalHtmlPasteScenario(page, context);
});
