import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

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
    const w = window as unknown as {
      __fcInst?: {
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          setNumber?: (addr: { sheet: number; row: number; col: number }, value: number) => void;
        };
        store?: {
          setState?: (fn: (state: Record<string, unknown>) => Record<string, unknown>) => void;
        };
        openPivotTableDialog?: () => void;
      };
    };
    const inst = w.__fcInst;
    if (!inst?.workbook || !inst.store?.setState || !inst.openPivotTableDialog) {
      throw new Error('window.__fcInst with PivotTable APIs is required');
    }
    const headers = ['Region', 'Sales', 'Qty', 'Channel', 'Segment'];
    headers.forEach((value, col) => {
      inst.workbook?.setText?.({ sheet: 0, row: 0, col }, value);
    });
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 0 }, 'East');
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 1 }, 10);
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 2 }, 2);
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 3 }, 'Online');
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 4 }, 'Consumer');
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 0 }, 'West');
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 1 }, 20);
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 2 }, 4);
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 3 }, 'Retail');
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 4 }, 'Business');
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...(state.selection as Record<string, unknown>),
        active: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 4 },
      },
    }));
    inst.openPivotTableDialog();
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
  await expect(filterDialog.locator('[data-pivot-filter-condition="true"]')).toHaveValue('none');
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
    const w = window as unknown as {
      __fcInst?: {
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          setNumber?: (addr: { sheet: number; row: number; col: number }, value: number) => void;
        };
        store?: {
          setState?: (fn: (state: Record<string, unknown>) => Record<string, unknown>) => void;
        };
      };
    };
    const inst = w.__fcInst;
    if (!inst?.workbook || !inst.store?.setState) {
      throw new Error('window.__fcInst with workbook/store APIs is required');
    }
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 0 }, 'Region');
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 1 }, 'Sales');
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 0 }, 'East');
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 1 }, 10);
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 0 }, 'West');
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 1 }, 20);
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...(state.selection as Record<string, unknown>),
        active: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      },
    }));
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
  const sourcePicker = pivotDialog.locator('[data-range-picker-kind="pivot-source"]');
  const destinationPicker = pivotDialog.locator('[data-range-picker-kind="pivot-destination"]');
  await expect(sourcePicker).toBeVisible();
  await expect(destinationPicker).toBeVisible();
  await expect(pivotDialog.getByText('Choose the data that you want to analyze.')).toBeVisible();
  await expect(pivotDialog.getByLabel('Table/Range')).toHaveValue('Sheet1!$A$1:$B$3');
  await expect(pivotDialog.getByLabel('New worksheet')).toBeChecked();
  await expect(pivotDialog.getByLabel('Existing worksheet')).not.toBeChecked();
  await sourcePicker.click();
  await expect(sourcePicker).toHaveAttribute('aria-pressed', 'true');
  await expect(pivotDialog).toHaveClass(/fc-fmtdlg--range-picking/);
  await destinationPicker.click();
  await expect(sourcePicker).toHaveAttribute('aria-pressed', 'false');
  await expect(destinationPicker).toHaveAttribute('aria-pressed', 'true');
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
