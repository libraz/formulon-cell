import { test } from '@playwright/test';

import {
  runCreateTableEntryPointsScenario,
  runFormatAsTableGalleryScenario,
  runMacCreateTableScenario,
} from '../../../tests/e2e-shared/scenarios/create-table.js';

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
