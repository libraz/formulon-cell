import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { setWorkbookStructureProtected } from '../../../../src/commands/protection.js';
import { addrKey } from '../../../../src/engine/address.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { seedNumber, stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

describe('Spreadsheet.mountToolbar', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  it('inserts and deletes cells through the Home Cells dropdowns', () => {
    seedNumber(sheet, 1, 1, 10);
    seedNumber(sheet, 2, 1, 20);
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    expect(insertButton).toBeTruthy();
    insertButton?.click();
    expect(host.querySelectorAll('#menu-insert-cells .fc-tb__menu-item--iconic').length).toBe(4);
    expect(host.querySelector<HTMLButtonElement>('[data-cell-insert="sheet"]')?.disabled).toBe(
      false,
    );
    const insertCellsButton = host.querySelector<HTMLButtonElement>('[data-cell-insert="cells"]');
    expect(insertCellsButton).toBeTruthy();
    const insertEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(insertEvent, 'target', { value: insertCellsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(insertEvent)).toBe(true);
    const insertDialog = document.querySelector<HTMLElement>('.fc-cellshift');
    expect(insertDialog).toBeTruthy();
    insertDialog?.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({
      kind: 'number',
      value: 10,
    });

    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );
    expect(deleteButton).toBeTruthy();
    deleteButton?.click();
    expect(host.querySelectorAll('#menu-delete-cells .fc-tb__menu-item--iconic').length).toBe(6);
    const disabledDeleteSheet = host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
    expect(disabledDeleteSheet?.disabled).toBe(true);
    expect(disabledDeleteSheet?.dataset.menuDisabledReason).toBe(
      'A workbook must contain at least one visible sheet.',
    );
    const deleteCellsButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="cells"]');
    expect(deleteCellsButton).toBeTruthy();
    const deleteEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(deleteEvent, 'target', { value: deleteCellsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(deleteEvent)).toBe(true);
    const deleteDialog = document.querySelector<HTMLElement>('.fc-cellshift');
    expect(deleteDialog).toBeTruthy();
    deleteDialog?.querySelector<HTMLButtonElement>('.fc-cellshift__button--primary')?.click();
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'number',
      value: 10,
    });

    tb.dispose();
  });

  it('directly inserts and deletes whole rows and columns from generic Cells actions', () => {
    seedNumber(sheet, 2, 0, 10);
    seedNumber(sheet, 0, 2, 20);
    sheet.instance.history.clear();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );
    expect(insertButton).toBeTruthy();
    expect(deleteButton).toBeTruthy();

    const clickGenericInsert = (): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'insertRows', menuId: 'menu-insert-cells' },
        insertButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>('[data-cell-insert="cells"]');
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };
    const clickGenericDelete = (): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'deleteRows', menuId: 'menu-delete-cells' },
        deleteButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>('[data-cell-delete="cells"]');
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };

    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 1,
      c0: 0,
      r1: 1,
      c1: 16_383,
    });
    clickGenericInsert();
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    expect(sheet.instance.history.undo()).toBe(false);

    mutators.setCopyRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 1,
      c0: 0,
      r1: 1,
      c1: 16_383,
    });
    clickGenericDelete();
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    expect(sheet.instance.history.undo()).toBe(false);

    mutators.setCopyRange(sheet.instance.store, null);
    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 0,
      c0: 1,
      r1: 1_048_575,
      c1: 1,
    });
    clickGenericInsert();
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({
      kind: 'number',
      value: 20,
    });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 20,
    });
    expect(sheet.instance.history.undo()).toBe(false);

    mutators.setCopyRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 0,
      c0: 1,
      r1: 1_048_575,
      c1: 1,
    });
    clickGenericDelete();
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 20,
    });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 20,
    });
    expect(sheet.instance.history.undo()).toBe(false);

    tb.dispose();
  });

  it('routes a whole-row cut snapshot through generic Insert Cells and restores one undo', async () => {
    seedNumber(sheet, 1, 0, 10);
    sheet.instance.history.clear();
    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 1,
      c0: 0,
      r1: 1,
      c1: 16_383,
    });
    const clipboard = sheet.instance.clipboard;
    if (!clipboard) throw new Error('Expected mounted clipboard handle.');
    await clipboard.runShortcut('cut');
    expect(clipboard.getSnapshot()?.mode).toBe('cut');
    mutators.setRange(sheet.instance.store, {
      sheet: 0,
      r0: 3,
      c0: 0,
      r1: 3,
      c1: 0,
    });

    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    expect(insertButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'insertRows', menuId: 'menu-insert-cells' },
      insertButton as HTMLButtonElement,
    );
    const insertCellsButton = host.querySelector<HTMLButtonElement>('[data-cell-insert="cells"]');
    expect(insertCellsButton).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: insertCellsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);

    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'blank' });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'blank' });
    expect(sheet.instance.history.undo()).toBe(false);

    tb.dispose();
  });

  it('inserts and deletes rows, columns, and sheets through the Home Cells dropdowns', () => {
    seedNumber(sheet, 1, 0, 10);
    seedNumber(sheet, 2, 0, 20);
    seedNumber(sheet, 0, 1, 11);
    seedNumber(sheet, 0, 2, 22);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );
    expect(insertButton).toBeTruthy();
    expect(deleteButton).toBeTruthy();
    const clickInsert = (action: string): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'insertRows', menuId: 'menu-insert-cells' },
        insertButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-cell-insert="${action}"]`);
      expect(button).toBeTruthy();
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };
    const clickDelete = (action: string): void => {
      tb.dropdownsApi?.openDynamicRibbonDropdown(
        { command: 'deleteRows', menuId: 'menu-delete-cells' },
        deleteButton as HTMLButtonElement,
      );
      const button = host.querySelector<HTMLButtonElement>(`[data-cell-delete="${action}"]`);
      expect(button).toBeTruthy();
      expect(button?.disabled).toBe(false);
      const event = new MouseEvent('click', { bubbles: true });
      Object.defineProperty(event, 'target', { value: button });
      expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    };

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    clickInsert('rows');
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
    expect(sheet.workbook.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });
    clickDelete('rows');
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'number',
      value: 10,
    });

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    clickInsert('cols');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({
      kind: 'number',
      value: 11,
    });
    clickDelete('cols');
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({
      kind: 'number',
      value: 11,
    });

    const initialSheetCount = sheet.workbook.sheetCount;
    clickInsert('sheet');
    expect(sheet.workbook.sheetCount).toBe(initialSheetCount + 1);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteRows', menuId: 'menu-delete-cells' },
      deleteButton as HTMLButtonElement,
    );
    const deleteSheetButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
    expect(deleteSheetButton?.disabled).toBe(true);
    expect(deleteSheetButton?.getAttribute('aria-disabled')).toBe('true');
    expect(deleteSheetButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot remove sheets.',
    );

    tb.dispose();
  });

  it('explains workbook structure protection on disabled Insert/Delete Sheet menu items', () => {
    setWorkbookStructureProtected(sheet.instance.store, true);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const insertButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="insertRows"]',
    );
    const deleteButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"]',
    );

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'insertRows', menuId: 'menu-insert-cells' },
      insertButton as HTMLButtonElement,
    );
    const insertSheetButton = host.querySelector<HTMLButtonElement>('[data-cell-insert="sheet"]');
    expect(insertSheetButton?.disabled).toBe(true);
    expect(insertSheetButton?.dataset.menuDisabledReason).toBe('Workbook structure is protected.');

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteRows', menuId: 'menu-delete-cells' },
      deleteButton as HTMLButtonElement,
    );
    const deleteSheetButton = host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
    expect(deleteSheetButton?.disabled).toBe(true);
    expect(deleteSheetButton?.dataset.menuDisabledReason).toBe('Workbook structure is protected.');

    tb.dispose();
  });

  it('applies row, protection, and sheet-tab formatting through the Home Format dropdown', async () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const formatButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"]',
    );
    expect(formatButton).toBeTruthy();
    formatButton?.click();
    expect(host.querySelectorAll('#menu-format-cells .fc-tb__menu-item--iconic').length).toBe(18);
    expect(host.querySelectorAll('#menu-format-cells .fc-tb__color-swatch').length).toBe(14);
    const visibilityTrigger = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-format-submenu="visibility"]',
    );
    const tabColorTrigger = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-format-submenu="tabColor"]',
    );
    const visibilityPanel = host.querySelector<HTMLElement>('#menu-format-cells-visibility');
    const tabColorPanel = host.querySelector<HTMLElement>('#menu-format-cells-tabColor');
    expect(visibilityTrigger?.getAttribute('aria-controls')).toBe('menu-format-cells-visibility');
    expect(tabColorTrigger?.getAttribute('aria-controls')).toBe('menu-format-cells-tabColor');
    expect(visibilityPanel?.hidden).toBe(true);
    expect(tabColorPanel?.hidden).toBe(true);
    visibilityTrigger?.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
    expect(visibilityPanel?.hidden).toBe(false);
    expect(visibilityTrigger?.getAttribute('aria-expanded')).toBe('true');
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.classList.contains('fc-tb__color-swatch'),
    ).toBe(true);
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-none"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.getAttribute('aria-checked'),
    ).toBe('false');
    tabColorTrigger?.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
    expect(tabColorPanel?.hidden).toBe(false);
    expect(visibilityPanel?.hidden).toBe(true);
    const lockCellButton = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-cell-format="lock-cell"]',
    );
    expect(lockCellButton?.getAttribute('role')).toBe('menuitemcheckbox');
    expect(lockCellButton?.getAttribute('aria-checked')).toBe('true');
    expect(lockCellButton?.classList.contains('fc-tb__menu-item--checked')).toBe(true);
    const showRowsBeforeHide = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="show-rows"]',
    );
    const showColsBeforeHide = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="show-cols"]',
    );
    const moveCopyButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="move-sheet-copy"]',
    );
    const renameSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="rename-sheet"]',
    );
    const unhideSheetButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="unhide-sheet"]',
    );
    expect(showRowsBeforeHide?.getAttribute('aria-disabled')).toBe('true');
    expect(showRowsBeforeHide?.dataset.menuDisabledReason).toBe('No hidden rows are selected.');
    expect(showColsBeforeHide?.getAttribute('aria-disabled')).toBe('true');
    expect(showColsBeforeHide?.dataset.menuDisabledReason).toBe('No hidden columns are selected.');
    expect(renameSheetButton?.getAttribute('aria-disabled')).toBe('true');
    expect(renameSheetButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot rename, move, hide, or unhide sheets.',
    );
    expect(moveCopyButton?.getAttribute('aria-disabled')).toBe('true');
    expect(moveCopyButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot rename, move, hide, or unhide sheets.',
    );
    expect(unhideSheetButton?.getAttribute('aria-disabled')).toBe('true');
    expect(unhideSheetButton?.dataset.menuDisabledReason).toBe(
      'This workbook engine cannot rename, move, hide, or unhide sheets.',
    );
    const hideRowsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="hide-rows"]');
    expect(hideRowsButton).toBeTruthy();
    const hideEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hideEvent, 'target', { value: hideRowsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hideEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(true);

    formatButton?.click();
    const showRowsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="show-rows"]');
    expect(showRowsButton).toBeTruthy();
    expect(showRowsButton?.getAttribute('aria-disabled')).toBe('false');
    expect(showRowsButton?.dataset.menuDisabledReason).toBeUndefined();
    const showEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(showEvent, 'target', { value: showRowsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(showEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenRows.has(1)).toBe(false);

    formatButton?.click();
    const unlockButton = host.querySelector<HTMLButtonElement>('[data-cell-format="unlock-cell"]');
    expect(unlockButton).toBeTruthy();
    const unlockEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unlockEvent, 'target', { value: unlockButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unlockEvent)).toBe(true);
    expect(
      sheet.instance.store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))
        ?.locked,
    ).toBe(false);
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'formatCellsHome', menuId: 'menu-format-cells' },
      formatButton as HTMLButtonElement,
    );
    const unlockedLockCellButton = host.querySelector<HTMLButtonElement>(
      '#menu-format-cells [data-cell-format="lock-cell"]',
    );
    expect(unlockedLockCellButton?.getAttribute('aria-checked')).toBe('false');
    expect(unlockedLockCellButton?.classList.contains('fc-tb__menu-item--checked')).toBe(false);

    formatButton?.click();
    const tabColorButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="tab-color-red"]',
    );
    expect(tabColorButton).toBeTruthy();
    const tabColorEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(tabColorEvent, 'target', { value: tabColorButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(tabColorEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
    formatButton?.click();
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.getAttribute('aria-checked'),
    ).toBe('true');
    expect(
      host
        .querySelector<HTMLButtonElement>('#menu-format-cells [data-cell-format="tab-color-red"]')
        ?.classList.contains('fc-tb__color-swatch--active'),
    ).toBe(true);

    formatButton?.click();
    const rowHeightButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="row-height"]',
    );
    expect(rowHeightButton).toBeTruthy();
    const rowHeightEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(rowHeightEvent, 'target', { value: rowHeightButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(rowHeightEvent)).toBe(true);
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const input = dialog?.querySelector<HTMLInputElement>('input');
    expect(input).toBeTruthy();
    if (!input) throw new Error('Row height prompt input was not rendered');
    input.value = '48';
    dialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();
    expect(sheet.instance.store.getState().layout.rowHeights.get(1)).toBe(48);

    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 });
    formatButton?.click();
    const hideColsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="hide-cols"]');
    expect(hideColsButton).toBeTruthy();
    const hideColsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(hideColsEvent, 'target', { value: hideColsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(hideColsEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenCols.has(1)).toBe(true);

    formatButton?.click();
    const showColsButton = host.querySelector<HTMLButtonElement>('[data-cell-format="show-cols"]');
    expect(showColsButton).toBeTruthy();
    expect(showColsButton?.getAttribute('aria-disabled')).toBe('false');
    expect(showColsButton?.dataset.menuDisabledReason).toBeUndefined();
    const showColsEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(showColsEvent, 'target', { value: showColsButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(showColsEvent)).toBe(true);
    expect(sheet.instance.store.getState().layout.hiddenCols.has(1)).toBe(false);

    formatButton?.click();
    const colWidthButton = host.querySelector<HTMLButtonElement>('[data-cell-format="col-width"]');
    expect(colWidthButton).toBeTruthy();
    const colWidthEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(colWidthEvent, 'target', { value: colWidthButton });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(colWidthEvent)).toBe(true);
    const colDialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    const colInput = colDialog?.querySelector<HTMLInputElement>('input');
    expect(colInput).toBeTruthy();
    if (!colInput) throw new Error('Column width prompt input was not rendered');
    colInput.value = '96';
    colDialog?.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Enter' }));
    await Promise.resolve();
    await Promise.resolve();
    expect(sheet.instance.store.getState().layout.colWidths.get(1)).toBe(96);

    tb.dispose();
  });

  it('disables the sheet move-or-copy entry when the host cannot reorder sheets', () => {
    const added = sheet.workbook.addSheet();
    expect(added).toBeGreaterThan(0);
    mutators.setSheetIndex(sheet.instance.store, added);
    mutators.setRange(sheet.instance.store, { sheet: added, r0: 0, c0: 0, r1: 0, c1: 0 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const formatButton = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatCellsHome"]',
    );
    expect(formatButton).toBeTruthy();
    formatButton?.click();
    const moveCopyButton = host.querySelector<HTMLButtonElement>(
      '[data-cell-format="move-sheet-copy"]',
    );
    expect(moveCopyButton).toBeTruthy();
    expect(moveCopyButton?.disabled).toBe(true);
    expect(moveCopyButton?.getAttribute('aria-disabled')).toBe('true');

    tb.dispose();
  });

  it('keeps Merge Cells as an icon split button and routes secondary merge actions', () => {
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    const helpers = stubHelpers();
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: {
        ...helpers,
        createIcon: (name) => {
          const icon = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          icon.setAttribute('class', 'fc-tb__rb-icon');
          icon.dataset.icon = name;
          return icon;
        },
      },
    });

    const mergeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(mergeButton).toBeTruthy();
    expect(mergeButton?.dataset.ribbonActivation).toBe('splitPrimary');
    expect(mergeButton?.dataset.ribbonMenuId).toBe('menu-merge');
    expect(mergeButton?.getAttribute('aria-haspopup')).toBe('menu');
    expect(mergeButton?.querySelector('.fc-tb__rb-icon')).toBeTruthy();
    const textLabels = Array.from(mergeButton?.querySelectorAll('span') ?? []).filter(
      (span) =>
        !span.classList.contains('fc-tb__rb-icon') &&
        !span.classList.contains('fc-tb__rb-split-chevron'),
    );
    expect(textLabels).toEqual([]);

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'merge', menuId: 'menu-merge' },
      mergeButton as HTMLButtonElement,
    );
    const mergeItems = Array.from(
      host.querySelectorAll<HTMLButtonElement>('#menu-merge .fc-tb__menu-item--iconic'),
    );
    expect(mergeItems.map((item) => item.textContent)).toEqual([
      'Merge & Center',
      'Merge Across',
      'Merge cells',
      'Unmerge Cells',
    ]);

    const mergeCenter = host.querySelector<HTMLButtonElement>(
      '#menu-merge [data-merge-action="mergeCenter"]',
    );
    const mergeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(mergeEvent, 'target', { value: mergeCenter });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(mergeEvent)).toBe(true);
    expect(
      sheet.instance.store.getState().merges.byAnchor.get(addrKey({ sheet: 0, row: 0, col: 0 })),
    ).toEqual({
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });

    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'merge', menuId: 'menu-merge' },
      mergeButton as HTMLButtonElement,
    );
    const unmerge = host.querySelector<HTMLButtonElement>(
      '#menu-merge [data-merge-action="unmergeCells"]',
    );
    const unmergeEvent = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(unmergeEvent, 'target', { value: unmerge });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(unmergeEvent)).toBe(true);
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);

    tb.dispose();
  });

  it('skips huge Merge Across selections before iterating each row', () => {
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 1 });
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });

    const mergeButton = host.querySelector<HTMLButtonElement>('[data-ribbon-command="merge"]');
    expect(mergeButton).toBeTruthy();
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'merge', menuId: 'menu-merge' },
      mergeButton as HTMLButtonElement,
    );
    const mergeAcross = host.querySelector<HTMLButtonElement>(
      '#menu-merge [data-merge-action="mergeAcross"]',
    );
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: mergeAcross });

    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    expect(sheet.instance.store.getState().merges.byAnchor.size).toBe(0);

    tb.dispose();
  });
});
