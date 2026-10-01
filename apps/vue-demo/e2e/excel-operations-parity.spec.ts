import { test } from '@playwright/test';
import {
  runCellBandReferencesScenario,
  runCrossSheetCutReferencesScenario,
  runDirectionalFillHistoryScenario,
  runMergedStructureCopyScenario,
  runPartialRangeCutScenario,
  runRepeatedPasteScenario,
  runSortFormulaHistoryScenario,
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

test('Excel parity: Fill Down copies text and undoes values and formats together', async ({
  page,
}) => {
  await runDirectionalFillHistoryScenario(page);
});

test('Excel parity: sorting adjusts formulas and undoes row formats', async ({ page }) => {
  await runSortFormulaHistoryScenario(page);
});

test('Excel parity: cell-band insertion and deletion update cross-sheet references', async ({
  page,
}) => {
  await runCellBandReferencesScenario(page);
});

test('Excel parity: repeated paste fills the selected range with formulas and formats', async ({
  page,
}) => {
  await runRepeatedPasteScenario(page);
});
