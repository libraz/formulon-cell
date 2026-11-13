import { test } from '@playwright/test';

import {
  runCreateNamesScenario,
  runCreateNamesTooLargeScenario,
} from '../../../tests/e2e-shared/scenarios/create-names.js';

for (const platform of ['default', 'mac'] as const) {
  test(`Create Names from Selection names each column after its header (${platform} chrome)`, async ({
    page,
  }) => {
    await runCreateNamesScenario(page, platform);
  });

  for (const locale of ['en', 'ja'] as const) {
    test(`Create Names from Selection over two whole rows says the selection is too large in ${locale} (${platform} chrome)`, async ({
      page,
    }) => {
      await runCreateNamesTooLargeScenario(page, locale, platform);
    });
  }
}
