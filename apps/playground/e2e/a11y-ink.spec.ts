import { test } from '@playwright/test';

import { runInkA11yScenario } from '../../../tests/e2e-shared/scenarios/a11y-ink.js';

test('A04 (playground): ink theme passes WCAG 2.2 AA (contrast included)', async ({
  page,
}, testInfo) => {
  await runInkA11yScenario(page, testInfo);
});
