import {
  runCreateTableEntryPointsScenario,
  runFormatAsTableGalleryScenario,
  runMacAccessibilityMenuScenario,
  runMacCreateTableScenario,
} from '../../../tests/e2e-shared/scenarios/create-table.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Ctrl+T, Ctrl+L and Insert > Table open Create Table on the current region', async ({
  page,
}) => {
  await runCreateTableEntryPointsScenario(page);
});

test('Format as Table opens the style gallery, then Create Table', async ({ page }) => {
  await runFormatAsTableGalleryScenario(page);
});

test('Mac Insert > Table proposes the region as a formula reference', async ({ page }) => {
  await runMacCreateTableScenario(page);
});

test('Mac Review > Check Accessibility formats the region as a table and unmerges cells', async ({
  page,
}) => {
  await runMacAccessibilityMenuScenario(page);
});
