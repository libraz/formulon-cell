import { test } from '@playwright/test';

import {
  runMacAuditMenusScenario,
  runMacAutoSumMenuScenario,
  runMacChartKindScenario,
  runMacCollapsedRibbonScenario,
  runMacConsolidateScenario,
  runMacDirectAutomationScenario,
  runMacDirectDataControlsScenario,
  runMacEditingPaletteScenario,
  runMacFormulaAndAutomationScenario,
  runMacFormulaPaletteScenario,
  runMacFunctionCatalogScenario,
  runMacGoalSeekScenario,
  runMacHomeMenusScenario,
  runMacInkScenario,
  runMacInsertGalleryScenario,
  runMacOutlineDetailScenario,
  runMacPageAndCalculationScenario,
  runMacPageToggleAndZoomScenario,
  runMacRibbonSwitchingScenario,
  runMacScaleAndSheetViewScenario,
  runMacSparklineScenario,
  runMacSubtotalScenario,
  runMacValidationMenuScenario,
  runMacViewShowScenario,
} from '../../../tests/e2e-shared/scenarios/mac-ribbon.js';

for (const locale of ['en', 'ja'] as const) {
  test(`Mac: ${locale} editing palette preserves drafts and edits a compound formula in one undo`, async ({
    page,
  }) => {
    await runMacEditingPaletteScenario(page, locale);
  });
}

test('Mac: all ribbon tabs switch their command groups with mouse and keyboard', async ({
  page,
}) => {
  await runMacRibbonSwitchingScenario(page);
});

test('Mac: Insert Sparkline creates an undoable sparkline', async ({ page }) => {
  await runMacSparklineScenario(page);
});

test('Mac: View Show menu toggles the formula bar', async ({ page }) => {
  await runMacViewShowScenario(page);
});

test('Mac: Draw creates and erases undoable ink, then restores cell editing', async ({ page }) => {
  await runMacInkScenario(page);
});

test('Mac: What-If Goal Seek previews, commits, cancels, and undoes', async ({ page }) => {
  await runMacGoalSeekScenario(page);
});

test('Mac: Data Consolidate writes one undoable result', async ({ page }) => {
  await runMacConsolidateScenario(page);
});

test('Mac: Page presets, breaks and calculation options execute their selected action', async ({
  page,
}) => {
  await runMacPageAndCalculationScenario(page);
});
test('Mac: Chart galleries preserve selected chart kinds and history', async ({ page }) => {
  await runMacChartKindScenario(page);
});
test('Mac: Subtotal inserts calculated groups in one undoable operation', async ({ page }) => {
  await runMacSubtotalScenario(page);
});

test('Mac: Function families calculate and automation shows its result', async ({ page }) => {
  await runMacFormulaAndAutomationScenario(page);
});

test('Mac: Home dropdowns retain formatting actions and history', async ({ page }) => {
  await runMacHomeMenusScenario(page);
});

test('Mac: collapsed ribbon tabs reveal commands and retain the display preference', async ({
  page,
}) => {
  await runMacCollapsedRibbonScenario(page);
});

test('Mac: function menus expose supported families and real recent history', async ({ page }) => {
  await runMacFunctionCatalogScenario(page);
});
test('Mac: outline detail controls collapse and expand rows and columns with undo', async ({
  page,
}) => {
  await runMacOutlineDetailScenario(page);
});

test('Mac: print scaling and saved sheet views execute from their ribbon groups', async ({
  page,
}) => {
  await runMacScaleAndSheetViewScenario(page);
});

test('Mac: Data exposes direct sorting and filter controls with real results', async ({ page }) => {
  await runMacDirectDataControlsScenario(page);
});
test('Mac: page option toggles retain checked state and View restores 100% zoom', async ({
  page,
}) => {
  await runMacPageToggleAndZoomScenario(page);
});

test('Mac: formula audit menus remove specific arrows and ignore actual errors', async ({
  page,
}) => {
  await runMacAuditMenusScenario(page);
});
test('Mac: Data Validation menu configures, circles, clears and restores rules', async ({
  page,
}) => {
  await runMacValidationMenuScenario(page);
});
test('Mac: automation scripts are directly available in the ribbon gallery', async ({ page }) => {
  await runMacDirectAutomationScenario(page);
});

test('Mac: supported shapes have a visual gallery and undoable insertion', async ({ page }) => {
  await runMacInsertGalleryScenario(page);
});

test('Mac: AutoSum menu computes a chosen aggregate and supports undo', async ({ page }) => {
  await runMacAutoSumMenuScenario(page);
});

test('Mac: Japanese ribbon labels and icons fit all tabs at four window widths', async ({
  page,
}) => {
  await runMacRibbonSwitchingScenario(page, 'ja');
});

test('Mac: formula palette previews without writes, retains Done, and reseeds from the ribbon', async ({
  page,
}) => {
  await runMacFormulaPaletteScenario(page);
});
