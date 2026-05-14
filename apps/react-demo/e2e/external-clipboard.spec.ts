import { test } from '@playwright/test';

import { runExternalHtmlPasteScenario } from '../../../tests/e2e-shared/scenarios/external-clipboard.js';

test('C03 (react-demo): external HTML payload pastes via system clipboard', async ({
  page,
  context,
  browserName,
}) => {
  test.skip(browserName === 'webkit', 'WebKit denies clipboard.write without a user gesture');
  await runExternalHtmlPasteScenario(page, context);
});
