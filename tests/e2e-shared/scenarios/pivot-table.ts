import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

/** P01 — a text row field's Filter > Label Filters... opens the panel, and a label condition there
 * opens the desktop Label Filter dialog. This goes through the demo-exposed core instance instead
 * of wrapper-specific APIs so React/Vue run the same scenario. */
export async function runPivotTableFilterDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const w = window as unknown as { __fcInst?: SpreadsheetInstance; __fcProbes: E2eProbes };
    const inst = w.__fcInst;
    if (!inst) throw new Error('window.__fcInst is required');
    const sheet = inst.workbook.sheet(0);
    inst.workbook.batch(() => {
      ['Channel', 'Sales'].forEach((value, col) => {
        sheet.setText(0, col, value);
      });
      sheet.setText(1, 0, 'Online');
      sheet.setNumber(1, 1, 10);
      sheet.setText(2, 0, 'Retail');
      sheet.setNumber(2, 1, 20);
    });
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(inst.workbook);
    const made = w.__fcProbes['commands/pivot-table'].createPivotTableFromRange(handle, {
      source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      destination: { sheet: 0, row: 0, col: 3 },
      name: 'ChannelReport',
      rowField: 'Channel',
      valueField: 'Sales',
      aggregation: 0,
    });
    if (!made.ok) throw new Error(`pivot creation failed: ${made.reason}`);
    inst.history.clear();
  });

  const label = await gridCellCenter(page, 1, 3);
  await page.mouse.click(label.x, label.y, { button: 'right' });
  await page.locator('.fc-ctxmenu [data-fc-action="filter"]').first().hover();
  await page.locator('.fc-ctxmenu__sub [data-fc-action="pivotFilterLabel"]').click();
  const panel = page.locator('.fc-pivot-filter-panel');
  await expect(panel).toBeVisible();
  // 12 is PivotFilterKind.LabelEqual.
  await panel.locator('select[data-pivot-filter-by="label"]').selectOption('12');

  const filterDialog = page.getByRole('dialog', { name: 'Label Filter (Channel)' });
  await expect(filterDialog).toBeVisible();
  // The panel closes as it hands over to the dialog.
  await expect(panel).toBeHidden();
  await expect(
    filterDialog.locator('[data-pivot-filter-operator="true"] option:checked'),
  ).toHaveText('equals');
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

/** PivotTable Options edits the report and cache options it shows; a reopen and an XLSX round
 *  trip both show the edited values. */
export async function runPivotTableOptionsScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();
  await sp.mount();
  await sp.expectNoStub();

  const buildReport = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst?: SpreadsheetInstance; __fcProbes: E2eProbes };
      const inst = w.__fcInst;
      if (!inst) throw new Error('window.__fcInst is required');
      const sheet = inst.workbook.sheet(0);
      inst.workbook.batch(() => {
        sheet.setText(0, 0, 'Channel');
        sheet.setText(0, 1, 'Sales');
        sheet.setText(1, 0, 'Online');
        sheet.setNumber(1, 1, 10);
        sheet.setText(2, 0, 'Retail');
        sheet.setNumber(2, 1, 20);
      });
      const handle = w.__fcProbes['api/workbook'].workbookHandleOf(inst.workbook);
      const made = w.__fcProbes['commands/pivot-table'].createPivotTableFromRange(handle, {
        source: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
        destination: { sheet: 0, row: 0, col: 3 },
        name: 'ChannelReport',
        rowField: 'Channel',
        valueField: 'Sales',
        aggregation: 0,
      });
      if (!made.ok) throw new Error(`pivot creation failed: ${made.reason}`);
      inst.history.clear();
    });
  const readOptions = () =>
    page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
      const o = handle.pivotOptions(0, 0);
      const c = handle.pivotCacheOptions(0, 0);
      if (!o || !c) return null;
      return {
        showDrill: o.showDrill,
        showError: o.showError,
        errorCaption: o.errorCaption,
        mergeItem: o.mergeItem,
        indent: o.indent,
        pageOverThenDown: o.pageOverThenDown,
        pageWrap: o.pageWrap,
        altText: o.altText,
        refreshOnLoad: c.refreshOnLoad,
      };
    });
  const dialog = page.getByRole('dialog', { name: 'PivotTable Options' });
  const tab = (id: string) => dialog.locator(`[data-pvopt-tab="${id}"]`);
  const option = (key: string) => dialog.locator(`[data-pvopt-option="${key}"]`);
  const openDialog = async () => {
    const cell = await gridCellCenter(page, 1, 3);
    await page.mouse.click(cell.x, cell.y, { button: 'right' });
    await page.locator('.fc-ctxmenu [data-fc-action="pivotOptions"]').click();
    await expect(dialog).toBeVisible();
  };
  const edited = {
    showDrill: false,
    showError: true,
    errorCaption: 'n/a',
    mergeItem: true,
    indent: 2,
    pageOverThenDown: true,
    pageWrap: 2,
    altText: 'Sales by channel',
    refreshOnLoad: true,
  };
  const expectEditedControls = async () => {
    await tab('display').click();
    await expect(option('showDrill')).not.toBeChecked();
    await expect(option('showError')).toBeChecked();
    await expect(
      dialog.locator('input[type="text"][data-pvopt-option="errorCaption"]'),
    ).toHaveValue('n/a');
    await tab('layout').click();
    await expect(option('mergeItem')).toBeChecked();
    // The popup shows each stored indent plus one.
    await expect(option('indent')).toHaveValue('3');
    await expect(option('pageOverThenDown').locator('input[value="1"]')).toBeChecked();
    await expect(option('pageWrap')).toHaveValue('2');
    await tab('data').click();
    await expect(dialog.locator('[data-pvopt-cache="refreshOnLoad"]')).toBeChecked();
    await tab('altText').click();
    await expect(option('altText')).toHaveValue('Sales by channel');
  };

  await buildReport();
  const before = await readOptions();
  if (!before) throw new Error('pivot options unreadable');
  // Every edited value differs from the new report's, so each one proves its control is bound.
  for (const [key, value] of Object.entries(edited))
    expect(before[key as keyof typeof edited], key).not.toEqual(value);
  await openDialog();

  await tab('display').click();
  await option('showDrill').uncheck();
  const caption = dialog.locator('input[type="text"][data-pvopt-option="errorCaption"]');
  // A caption applies only while its check box is on.
  await expect(caption).toBeDisabled();
  await option('showError').check();
  await expect(caption).toBeEnabled();
  await caption.fill('n/a');
  await tab('layout').click();
  await option('mergeItem').check();
  await option('indent').selectOption('3');
  await option('pageOverThenDown').locator('input[value="1"]').check();
  await option('pageWrap').selectOption('2');
  await tab('data').click();
  await dialog.locator('[data-pvopt-cache="refreshOnLoad"]').check();
  await tab('altText').click();
  await option('altText').fill('Sales by channel');
  await dialog.locator('[data-pvopt-ok]').click();
  await expect(dialog).toBeHidden();
  await expect.poll(readOptions).toEqual(edited);

  // Cancel keeps every value.
  await openDialog();
  await expectEditedControls();
  await tab('layout').click();
  await option('mergeItem').uncheck();
  await dialog.locator('[data-pvopt-cancel]').click();
  await expect(dialog).toBeHidden();
  expect(await readOptions()).toEqual(edited);

  await page.evaluate(async () => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { Workbook } = w.__fcProbes['api/workbook'];
    const next = await Workbook.load(w.__fcInst.workbook.save().bytes);
    if (next.engine !== 'wasm') throw new Error('Reload must use the real WASM engine');
    w.__fcInst.update({ workbook: next });
  });
  await expect.poll(readOptions).toEqual(edited);
  await openDialog();
  await expectEditedControls();
  await dialog.locator('[data-pvopt-cancel]').click();
  await expect(dialog).toBeHidden();

  expect(consoleErrors.read()).toEqual([]);
}
