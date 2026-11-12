import { test } from '@playwright/test';

import { runCreateNamesTooLargeScenario } from '../../../tests/e2e-shared/scenarios/create-names.js';

for (const platform of ['default', 'mac'] as const) {
  for (const locale of ['en', 'ja'] as const) {
    test(`Create Names from Selection over two whole rows says the selection is too large in ${locale} (${platform} chrome)`, async ({
      page,
    }) => {
      await runCreateNamesTooLargeScenario(page, locale, platform);
    });
  }
}
