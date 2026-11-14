import {
  runStaleHoverTipScenario,
  runStaleValueSaveScenario,
  runStatusBarRecalcScenario,
} from '../../../tests/e2e-shared/scenarios/stale-values.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  for (const recalcBeforeSave of [true, false]) {
    test(`stale values through an .xlsx save, recalculate before saving ${recalcBeforeSave ? 'on' : 'off'} (${platform} chrome)`, async ({
      page,
    }) => {
      await runStaleValueSaveScenario(page, platform, recalcBeforeSave);
    });
  }
}

for (const [platform, locale] of [
  ['mac', 'en'],
  ['mac', 'ja'],
  ['default', 'en'],
] as const) {
  test(`status bar Calculate recalculates stale values (${platform} chrome, ${locale})`, async ({
    page,
  }) => {
    await runStatusBarRecalcScenario(page, platform, locale);
  });
}

for (const platform of ['default', 'mac'] as const) {
  test(`hovering a stale value's triangle shows the Stale Value tip (${platform} chrome)`, async ({
    page,
  }) => {
    await runStaleHoverTipScenario(page, platform);
  });
}
