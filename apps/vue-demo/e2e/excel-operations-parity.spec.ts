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
import { SYSTEM_CLIPBOARD } from '../../../tests/e2e-shared/tags.js';
import { test } from '../../../tests/e2e-shared/test.js';

test('Excel parity: merged row edits, formatted copy, undo and redo', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runMergedStructureCopyScenario(page);
});

test('Excel parity: cross-sheet cut preserves and follows references', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runCrossSheetCutReferencesScenario(page);
});

test('Excel parity: partial range cut updates only source-sheet references', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
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

test('Excel parity: repeated paste fills the selected range with formulas and formats', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runRepeatedPasteScenario(page);
});

test('Excel parity: whole-column copy and insert preserve row positions', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runWholeColumnCopyInsertScenario(page);
});

test('Excel parity: whole-row copy and insert preserve column positions', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runWholeRowCopyInsertScenario(page);
});

test('Excel parity: whole-band keyboard and ribbon insertion and deletion', async ({ page }) => {
  await runWholeBandInsertDeleteScenario(page);
});

test('Excel parity: copied-band insert routes refresh references before the source', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runCopiedBandInsertRoutesScenario(page);
});

test('Excel parity: cut columns and rows insert through keyboard and ribbon', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runCutBandInsertRoutesScenario(page);
});

test('Excel parity: cross-sheet cut column insertion leaves the source column empty', {
  tag: SYSTEM_CLIPBOARD,
}, async ({ page }) => {
  await runCrossSheetCutBandInsertScenario(page);
});
