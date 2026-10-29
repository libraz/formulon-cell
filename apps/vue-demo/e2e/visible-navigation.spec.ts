import { test } from '@playwright/test';
import {
  runHiddenBandCommit,
  runHiddenBandNavigation,
  runHiddenBandTraversal,
  runHiddenGapMergeNavigation,
  runHiddenMergeNavigation,
  runMergedSelectionCycles,
} from '../../../tests/e2e-shared/scenarios/visible-navigation.js';

for (const platform of ['mac', 'default'] as const) {
  test(`Excel: ${platform} keyboard skips hidden rows and columns`, async ({ page }) => {
    await runHiddenBandNavigation(page, platform);
  });
}
test('Excel: inline and formula-bar commits move to visible cells', async ({ page }) => {
  await runHiddenBandCommit(page);
});
test('Excel: selection traversal skips hidden bands in both directions', async ({ page }) => {
  await runHiddenBandTraversal(page);
});

test('Excel: hidden merge edges match directional native navigation', async ({ page }) => {
  await runHiddenMergeNavigation(page);
});

test('Excel: merged selections complete forward and reverse cycles', async ({ page }) => {
  await runMergedSelectionCycles(page);
});
test('Excel: hidden gaps do not enter hidden leading merge edges', async ({ page }) => {
  await runHiddenGapMergeNavigation(page);
});
