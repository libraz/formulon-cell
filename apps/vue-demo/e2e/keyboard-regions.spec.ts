import {
  runGridOnlyRegionScenario,
  runModalKeyOwnershipScenario,
  runRegionCycleScenario,
  runSheetFocusTargetScenario,
} from '../../../tests/e2e-shared/scenarios/keyboard-regions.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const platform of ['default', 'mac'] as const) {
  for (const variant of ['none', 'watch', 'pivotFields', 'autoHide', 'noRibbon'] as const) {
    test(`F6 (vue-demo, ${platform}, ${variant}): F6 and Shift+F6 cycle the regions and open panes`, async ({
      page,
    }) => {
      await runRegionCycleScenario(page, platform, variant);
    });
  }

  test(`F6 (vue-demo, ${platform}): F6 is left unhandled when the grid is the only region`, async ({
    page,
  }) => {
    await runGridOnlyRegionScenario(page, platform);
  });

  test(`K01 (vue-demo, ${platform}): a modal dialog owns the region, ribbon and search keys`, async ({
    page,
  }) => {
    await runModalKeyOwnershipScenario(page, platform);
  });

  test(`K02 (vue-demo, ${platform}): the grid is the sheet's tab stop`, async ({ page }) => {
    await runSheetFocusTargetScenario(page, platform);
  });
}
