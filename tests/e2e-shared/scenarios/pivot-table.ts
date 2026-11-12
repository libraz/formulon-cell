import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** P01 — the shared PivotTable filter dialog is reachable from the
 * Create PivotTable Field Settings flow. This intentionally goes through the
 * demo-exposed core instance instead of wrapper-specific APIs so React/Vue run
 * the same scenario. */
export async function runPivotTableFilterDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      ['Region', 'Sales', 'Qty', 'Channel', 'Segment'].forEach((value, col) => {
        sheet.setText(0, col, value);
      });
      sheet.setText(1, 0, 'East');
      sheet.setNumber(1, 1, 10);
      sheet.setNumber(1, 2, 2);
      sheet.setText(1, 3, 'Online');
      sheet.setText(1, 4, 'Consumer');
      sheet.setText(2, 0, 'West');
      sheet.setNumber(2, 1, 20);
      sheet.setNumber(2, 2, 4);
      sheet.setText(2, 3, 'Retail');
      sheet.setText(2, 4, 'Business');
    });
    inst.history.clear();
    inst.selection.set('A1:E3');
    inst.dialogs.open('pivotTable');
  });

  const pivotDialog = page.getByRole('dialog', { name: 'Create PivotTable' });
  await expect(pivotDialog).toBeVisible();
  await pivotDialog.locator('[data-pivot-field-list-field="Segment"]').check();
  await pivotDialog.locator('[data-pivot-field-list-field="Channel"]').check();
  await pivotDialog.getByRole('button', { name: 'Field Settings: Channel' }).click();
  await expect(pivotDialog.locator('.fc-pivotdlg__area-settings-panel')).toContainText(
    'Field Settings: Channel',
  );
  await pivotDialog.getByRole('button', { name: 'Filter...' }).click();

  const filterDialog = page.getByRole('dialog', { name: 'PivotTable Filter: Channel' });
  await expect(filterDialog).toBeVisible();
  await expect(filterDialog.locator('[data-pivot-filter-category="true"]')).toHaveValue('label');
  // A new filter starts on the None condition.
  await expect(
    filterDialog.locator('[data-pivot-filter-condition="true"] option:checked'),
  ).toHaveText('None');
  await filterDialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(filterDialog).toBeHidden();

  expect(consoleErrors.read()).toEqual([]);
}

export async function runPivotTableRibbonPrimaryDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const inst = (window as unknown as { __fcInst?: SpreadsheetInstance }).__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      sheet.setText(0, 0, 'Region');
      sheet.setText(0, 1, 'Sales');
      sheet.setText(1, 0, 'East');
      sheet.setNumber(1, 1, 10);
      sheet.setText(2, 0, 'West');
      sheet.setNumber(2, 1, 20);
    });
    inst.history.clear();
    inst.selection.set('A1:B3');
  });

  await page.locator('[data-ribbon-tab="insert"]').click();
  await expect(page.locator('[data-ribbon-tab="insert"][aria-selected="true"]')).toHaveCount(1);

  const pivotButton = page.locator('[data-ribbon-command="pivotTableInsert"]').first();
  await expect(pivotButton).toBeVisible();
  await expect(pivotButton).toHaveAttribute('aria-haspopup', 'menu');
  await expect(pivotButton).toHaveAttribute('data-ribbon-activation', 'splitPrimary');
  await pivotButton.click();

  const pivotDialog = page.getByRole('dialog', { name: 'Create PivotTable' });
  await expect(pivotDialog).toBeVisible();
  await expect(page.locator('#menu-pivot-table')).toBeHidden();
  const sourcePicker = pivotDialog.locator('[data-range-picker="pivot-source"]');
  const destinationPicker = pivotDialog.locator('[data-range-picker="pivot-destination"]');
  await expect(sourcePicker).toBeVisible();
  await expect(destinationPicker).toBeVisible();
  await expect(pivotDialog.getByText('Choose the data that you want to analyze.')).toBeVisible();
  await expect(pivotDialog.getByLabel('Table/Range')).toHaveValue('Sheet1!$A$1:$B$3');
  await expect(pivotDialog.getByLabel('New worksheet')).toBeChecked();
  await expect(pivotDialog.getByLabel('Existing worksheet')).not.toBeChecked();
  // A location is only editable on an existing sheet, so the destination
  // picker stays disabled until that placement is selected.
  await expect(destinationPicker).toBeDisabled();
  await pivotDialog.getByLabel('Existing worksheet').check();
  await expect(destinationPicker).toBeEnabled();
  await sourcePicker.click();
  await expect(sourcePicker).toHaveAttribute('aria-pressed', 'true');
  await expect(pivotDialog).toHaveClass(/fc-fmtdlg--range-picking/);
  // Picking collapses the dialog down to the field being picked, so the source
  // pick has to end before the destination field is reachable again.
  await page.keyboard.press('Escape');
  await expect(sourcePicker).toHaveAttribute('aria-pressed', 'false');
  await destinationPicker.click();
  await expect(sourcePicker).toHaveAttribute('aria-pressed', 'false');
  await expect(destinationPicker).toHaveAttribute('aria-pressed', 'true');
  await expect(pivotDialog).toHaveClass(/fc-fmtdlg--range-picking/);
  const canvasBox = await page.locator('.fc-host__canvas').first().boundingBox();
  expect(canvasBox, 'canvas must be laid out').not.toBeNull();
  if (!canvasBox) throw new Error('canvas not measured');
  const passThrough = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      return {
        closestHost: !!el?.closest('.fc-host'),
        closestDialog: !!el?.closest('.fc-fmtdlg'),
      };
    },
    {
      x: Math.round(canvasBox.x + canvasBox.width / 2),
      y: Math.round(canvasBox.y + canvasBox.height / 2),
    },
  );
  expect(passThrough.closestHost).toBe(true);
  expect(passThrough.closestDialog).toBe(false);
  await page.keyboard.press('Escape');
  await expect(destinationPicker).toHaveAttribute('aria-pressed', 'false');
  await expect(pivotDialog).not.toHaveClass(/fc-fmtdlg--range-picking/);
  await pivotDialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(pivotDialog).toBeHidden();

  expect(consoleErrors.read()).toEqual([]);
}

/** P03 — on a protected sheet Insert PivotTable is unavailable from the ribbon and the instance
 *  cannot open its dialog; both come back once the sheet is unprotected. */
export async function runPivotTableProtectedSheetScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();
  const canOpenPivot = () =>
    page.evaluate(() => {
      const inst = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      return inst.dialogs.canOpen('pivotTable');
    });

  await page.evaluate(async () => {
    const inst = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const sheet = inst.workbook.sheet(0);
    sheet.setText(0, 0, 'Region');
    sheet.setText(0, 1, 'Sales');
    sheet.setText(1, 0, 'East');
    sheet.setNumber(1, 1, 10);
    inst.history.clear();
    inst.selection.set('A1:B2');
    await sheet.protect({ allowPivotTables: true });
  });

  await page.locator('[data-ribbon-tab="insert"]').click();
  const pivotButton = page.locator('[data-ribbon-command="pivotTableInsert"]').first();
  await expect(pivotButton).toBeDisabled();
  expect(await canOpenPivot()).toBe(false);

  await page.evaluate(() =>
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.workbook.sheet(0).unprotect(),
  );
  await expect(pivotButton).toBeEnabled();
  expect(await canOpenPivot()).toBe(true);
  expect(consoleErrors.read()).toEqual([]);
}
