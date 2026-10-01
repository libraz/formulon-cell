import { test } from '@playwright/test';
import {
  runCellBandReferencesScenario,
  runCopiedBandInsertRoutesScenario,
  runCrossSheetCutBandInsertScenario,
  runCrossSheetCutReferencesScenario,
  runCutBandInsertRoutesScenario,
  runDirectionalFillHistoryScenario,
  runMergedStructureCopyScenario,
  runPartialRangeCutScenario,
  runRepeatedPasteScenario,
  runSortFormulaHistoryScenario,
  runWholeBandInsertDeleteScenario,
  runWholeColumnCopyInsertScenario,
  runWholeRowCopyInsertScenario,
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

test('Excel parity: whole-column copy and insert preserve row positions', async ({ page }) => {
  await runWholeColumnCopyInsertScenario(page);
});

test('Excel parity: whole-row copy and insert preserve column positions', async ({ page }) => {
  await runWholeRowCopyInsertScenario(page);
});

test('Excel parity: whole-band keyboard and ribbon insertion and deletion', async ({ page }) => {
  await runWholeBandInsertDeleteScenario(page);
});

test('Excel parity: copied-band insert routes refresh references before the source', async ({
  page,
}) => {
  await runCopiedBandInsertRoutesScenario(page);
});

test('Excel parity: cut columns and rows insert through keyboard and ribbon', async ({ page }) => {
  await runCutBandInsertRoutesScenario(page);
});

test('Excel parity: cross-sheet cut column insertion leaves the source column empty', async ({
  page,
}) => {
  await runCrossSheetCutBandInsertScenario(page);
});
