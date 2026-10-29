import { test } from '@playwright/test';
import {
  runTabReturnCommits,
  runTabReturnMergesAndHidden,
  runTabReturnNavigation,
  runTabReturnResets,
} from '../../../tests/e2e-shared/scenarios/excel-entry-navigation.js';

test('Excel: Return after forward Tabs starts the next input row', async ({ page }) => {
  await runTabReturnNavigation(page);
});
test('Excel: inline and formula-bar entry share the Tab row origin', async ({ page }) => {
  await runTabReturnCommits(page);
});
test('Excel: Tab entry keeps its row across merges and hidden bands', async ({ page }) => {
  await runTabReturnMergesAndHidden(page);
});
test('Excel: reverse movement resets the Return origin and selections retain traversal', async ({
  page,
}) => {
  await runTabReturnResets(page);
});
