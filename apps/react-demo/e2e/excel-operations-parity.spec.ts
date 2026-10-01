import { test } from '@playwright/test';
import {
  runCrossSheetCutReferencesScenario,
  runMergedStructureCopyScenario,
  runPartialRangeCutScenario,
} from '../../../tests/e2e-shared/scenarios/excel-operations-parity.js';

test.skip(({ browserName }) => browserName === 'webkit', 'WebKit denies clipboard reads');

test('Excel parity: merged row edits, formatted copy, undo and redo', async ({ page }) => {
  await runMergedStructureCopyScenario(page);
});

test('Excel parity: cross-sheet cut preserves and follows references', async ({ page }) => {
  await runCrossSheetCutReferencesScenario(page);
});

test('Excel parity: partial range cut updates only source-sheet references', async ({ page }) => {
  await runPartialRangeCutScenario(page);
});
