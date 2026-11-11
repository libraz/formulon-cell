import { test } from '@playwright/test';

import { runSparklineProtectionScenario } from '../../../tests/e2e-shared/scenarios/sparkline-protection.js';

for (const platform of ['default', 'mac'] as const) {
  for (const cells of ['locked', 'unlocked'] as const) {
    test(`a protected sheet keeps the Sparkline tab but blocks every edit on ${cells} sparkline cells (${platform} chrome)`, async ({
      page,
    }) => {
      await runSparklineProtectionScenario(page, platform, cells);
    });
  }
}
