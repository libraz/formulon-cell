import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { canExecuteBuiltIn } from '../../../src/commands/built-in-command-policy.js';
import {
  addAllowedEditRange,
  removeAllowedEditRange,
  setWorkbookStructureProtected,
} from '../../../src/commands/protection.js';
import {
  copySheetAction,
  moveSheetAction,
  type SheetActionContext,
} from '../../../src/commands/sheet-actions.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createCellsDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/cells.js';
import { hydrateActiveSheetFromEngine } from '../../../src/mount/hydration.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('copy sheet', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    const workbook = await WorkbookHandle.createDefault();
    sheet = await mountStubSheet({ workbook, locale: 'en' });
    const wb = sheet.workbook;
    wb.addSheet('Second');
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 2);
    wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*10');
    wb.recalc();
    mutators.replaceCells(sheet.instance.store, wb.cells(0));
    mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    mutators.setSheetTabColor(sheet.instance.store, 0, '#ff0000');
    sheet.instance.history.clear();
  });
  afterEach(() => sheet.dispose());

  const ctx = (): SheetActionContext => ({
    store: sheet.instance.store,
    workbook: sheet.workbook,
    history: sheet.instance.history,
    origin: 'contextMenu',
  });
  const names = (): string[] =>
    Array.from({ length: sheet.workbook.sheetCount }, (_, i) => sheet.workbook.sheetName(i));
  const active = (): number => sheet.instance.store.getState().data.sheetIndex;

  it('copies the sheet before the chosen one under the engine name and activates it', () => {
    const copy = copySheetAction(ctx(), 0, 1);
    expect(copy).toBe(1);
    expect(names()).toEqual(['Sheet1', 'Sheet1 (2)', 'Second']);
    expect(active()).toBe(1);
    const wb = sheet.workbook;
    expect(wb.getValue({ sheet: 1, row: 0, col: 0 })).toEqual({ kind: 'number', value: 2 });
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 1 })).toBe('=A1*10');
    expect(wb.getValue({ sheet: 1, row: 0, col: 1 })).toEqual({ kind: 'number', value: 20 });
    hydrateActiveSheetFromEngine(wb, sheet.instance.store);
    const state = sheet.instance.store.getState();
    expect(state.format.formats.get(addrKey({ sheet: 1, row: 0, col: 0 }))?.bold).toBe(true);
    expect(state.layout.sheetTabColors.get(1)).toBe('#ff0000');
    expect(state.layout.sheetTabColors.get(0)).toBe('#ff0000');
    expect(state.layout.sheetTabColors.has(2)).toBe(false);
  });

  it('appends the copy at the end', () => {
    expect(copySheetAction(ctx(), 0, sheet.workbook.sheetCount)).toBe(2);
    expect(names()).toEqual(['Sheet1', 'Second', 'Sheet1 (2)']);
  });

  it('is one undo step that removes the copy and restores the active sheet', () => {
    const { history, store } = sheet.instance;
    copySheetAction(ctx(), 0, 0);
    expect(names()).toEqual(['Sheet1 (2)', 'Sheet1', 'Second']);
    expect(store.getState().layout.sheetTabColors.get(1)).toBe('#ff0000');
    expect(history.undo()).toBe(true);
    expect(names()).toEqual(['Sheet1', 'Second']);
    expect(active()).toBe(0);
    expect(store.getState().layout.sheetTabColors.get(0)).toBe('#ff0000');
    expect(store.getState().layout.sheetTabColors.has(1)).toBe(false);
    expect(history.canUndo()).toBe(false);
    expect(history.redo()).toBe(true);
    expect(names()).toEqual(['Sheet1 (2)', 'Sheet1', 'Second']);
    expect(active()).toBe(0);
  });

  it('is refused under workbook-structure protection and a policy denying sheet', () => {
    const { store } = sheet.instance;
    setWorkbookStructureProtected(store, true);
    expect(canExecuteBuiltIn(store, 'copySheet', 'contextMenu')).toMatchObject({
      allowed: false,
      code: 'protected',
    });
    expect(copySheetAction(ctx(), 0, 1)).toBe(-1);
    setWorkbookStructureProtected(store, false);
    sheet.instance.setPolicy({ defaultOperation: 'allow', operations: { sheet: false } });
    expect(canExecuteBuiltIn(store, 'copySheet', 'contextMenu').allowed).toBe(false);
    expect(copySheetAction(ctx(), 0, 1)).toBe(-1);
    expect(names()).toEqual(['Sheet1', 'Second']);
    expect(sheet.instance.history.canUndo()).toBe(false);
  });

  it('gives the copy its own allowed edit ranges, read from the engine', () => {
    const { store } = sheet.instance;
    const sourceId = addAllowedEditRange(
      store,
      { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 },
      {
        title: 'Inputs',
        extraAreas: [{ sheet: 0, r0: 5, c0: 0, r1: 5, c1: 0 }],
        workbook: sheet.workbook,
      },
    );
    const copy = copySheetAction(ctx(), 0, 2);
    const ranges = store.getState().protection.allowedEditRanges;
    const copied = ranges.find((entry) => entry.range.sheet === copy);
    expect(copied?.title).toBe('Inputs');
    expect(copied?.id).not.toBe(sourceId);
    expect(copied?.extraAreas?.map((area) => area.sheet)).toEqual([copy]);
    if (!copied) throw new Error('copy has no allowed range');
    removeAllowedEditRange(store, copied.id, { workbook: sheet.workbook });
    const left = store.getState().protection.allowedEditRanges;
    expect(left.map((entry) => entry.id)).toEqual([sourceId]);
    expect(sheet.workbook.getProtectedRanges(0).map((record) => record.name)).toEqual(['Inputs']);
    expect(sheet.workbook.getProtectedRanges(copy)).toEqual([]);
  });

  it('activates the restored sheet with its own layout on undo', () => {
    const { history, store } = sheet.instance;
    sheet.workbook.setColumnWidth(0, 0, 0, 30);
    sheet.workbook.setColumnWidth(1, 0, 0, 8);
    mutators.setSheetIndex(store, 1);
    hydrateActiveSheetFromEngine(sheet.workbook, store);
    const secondWidth = store.getState().layout.colWidths.get(0);
    copySheetAction(ctx(), 0, 2);
    const copyWidth = store.getState().layout.colWidths.get(0);
    expect(copyWidth).not.toBe(secondWidth);
    expect(history.undo()).toBe(true);
    expect(active()).toBe(1);
    expect(store.getState().layout.colWidths.get(0)).toBe(secondWidth);
  });

  it('keeps table overlays on their sheet through a move and its undo', () => {
    const { history, store } = sheet.instance;
    const wb = sheet.workbook;
    wb.setText({ sheet: 1, row: 0, col: 0 }, 'Head');
    wb.setNumber({ sheet: 1, row: 1, col: 0 }, 1);
    wb.engineWorkbook().sheet(1).addTable({ range: 'A1:A2', name: 'Scores' });
    copySheetAction(ctx(), 1, 2);
    expect(history.undo()).toBe(true);
    const sheetOf = (name: string) =>
      store.getState().tables.tables.find((table) => table.name === name)?.range.sheet;
    expect(sheetOf('Scores')).toBe(1);
    expect(moveSheetAction(ctx(), 1, 0)).toBe(true);
    expect(sheetOf('Scores')).toBe(0);
    expect(history.undo()).toBe(true);
    expect(sheetOf('Scores')).toBe(1);
    expect(history.redo()).toBe(true);
    expect(sheetOf('Scores')).toBe(0);
  });

  it('tells the user when the engine refuses the copy', async () => {
    sheet.workbook
      .engineWorkbook()
      .sheet(0)
      .addShape({ text: 'x', anchor: { fromRow: 5, fromCol: 5, toRow: 8, toCol: 8 } });
    const defaults = createCellsDropdownDefaults(sheet.instance, {});
    const pending = defaults.applyCellFormatAction('move-sheet-copy');
    await Promise.resolve();
    document.querySelector<HTMLInputElement>('.fc-movecopy__copy input')?.click();
    document.querySelector<HTMLButtonElement>('.fc-movecopy .fc-fmtdlg__btn--primary')?.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    const message = document.querySelector<HTMLElement>(
      '[role="alertdialog"] .fc-tb__dlg__message',
    );
    expect(message?.textContent).toContain("can't be copied");
    expect(names()).toEqual(['Sheet1', 'Second']);
    message
      ?.closest<HTMLElement>('[role="alertdialog"]')
      ?.querySelector<HTMLButtonElement>('button')
      ?.click();
    await pending;
  });

  it('copies through the Move or Copy dialog with Create a copy checked', async () => {
    const defaults = createCellsDropdownDefaults(sheet.instance, {});
    const pending = defaults.applyCellFormatAction('move-sheet-copy');
    await Promise.resolve();
    const copy = document.querySelector<HTMLInputElement>('.fc-movecopy__copy input');
    expect(copy?.disabled).toBe(false);
    copy?.click();
    const end = document.querySelectorAll<HTMLElement>('.fc-movecopy__list [role="option"]');
    end[end.length - 1]?.click();
    document.querySelector<HTMLButtonElement>('.fc-movecopy .fc-fmtdlg__btn--primary')?.click();
    await pending;
    expect(names()).toEqual(['Sheet1', 'Second', 'Sheet1 (2)']);
    expect(active()).toBe(2);
  });
});
