import { runMacRibbonFitScenario } from '../../../tests/e2e-shared/scenarios/mac-ribbon-fit.js';
import { test } from '../../../tests/e2e-shared/test.js';

for (const locale of ['fr', 'de', 'en', 'ja'] as const) {
  test(`Mac: ribbon tabs fit a 1440 wide window as measured (${locale})`, async ({ page }) => {
    // Pinned widths are macOS system-font text widths; other hosts fall back to other fonts.
    test.skip(process.platform !== 'darwin', 'Ribbon fit widths are pinned to macOS font metrics.');
    await runMacRibbonFitScenario(page, locale);
  });
}
