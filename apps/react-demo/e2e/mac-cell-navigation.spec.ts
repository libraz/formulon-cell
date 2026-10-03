import { test } from '@playwright/test';
import {
  runMacCollapseSelection,
  runMacCommandReturnFill,
  runMacFormulaBarPointerReferences,
  runMacFormulaBarSelection,
  runMacOptionSheetSwitch,
  runMacSelectionEntry,
  runMacSelectionTraversal,
} from '../../../tests/e2e-shared/scenarios/mac-cell-navigation.js';

test('Mac: Return and Tab cycle within the selected rectangle', async ({ page }) => {
  await runMacSelectionTraversal(page);
});
test('Mac: committing entries preserves and traverses a selection', async ({ page }) => {
  await runMacSelectionEntry(page);
});
test('Mac: Command Return fills the range, Option Return inserts a newline', async ({ page }) => {
  await runMacCommandReturnFill(page);
});
test('Mac: Shift Delete collapses the selection without clearing data', async ({ page }) => {
  await runMacCollapseSelection(page);
});

test('Mac: formula bar traversal and Command Return preserve selection and history', async ({
  page,
}) => {
  await runMacFormulaBarSelection(page);
});
test('Mac: formula bar accepts clicked and dragged range references', async ({ page }) => {
  await runMacFormulaBarPointerReferences(page);
});

test('Mac: Option arrow shortcuts switch worksheets', async ({ page }) => {
  await runMacOptionSheetSwitch(page);
});
