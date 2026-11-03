import { describe, expect, it } from 'vitest';
import {
  blockedByProtection,
  isCellWritable,
  setCellLocked,
  setProtectedSheet,
  verifySheetProtectionPassword,
} from '../../../src/commands/protection.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// Excel Worksheet.Protect results.
describe('sheet protection as Excel measures it', () => {
  const protectedStore = (password?: string) => {
    const store = createSpreadsheetStore();
    setCellLocked(store, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, false);
    setProtectedSheet(store, 0, true, password ? { password } : {});
    return store;
  };

  it('blocks locked cells and leaves unlocked ones writable', () => {
    const state = protectedStore().getState();
    expect(isCellWritable(state, { sheet: 0, row: 0, col: 0 })).toBe(false);
    expect(isCellWritable(state, { sheet: 0, row: 1, col: 1 })).toBe(true);
  });

  it('blocks row insertion and deletion by default', () => {
    const store = protectedStore();
    expect(blockedByProtection(store, 0, 'insertRows')).toBe(true);
    expect(blockedByProtection(store, 0, 'deleteRows')).toBe(true);
  });

  it('only the right password unprotects; none and a wrong one keep it protected', async () => {
    const state = protectedStore('pw1').getState();
    expect(await verifySheetProtectionPassword(state, 0, '')).toBe(false);
    expect(await verifySheetProtectionPassword(state, 0, 'x')).toBe(false);
    expect(await verifySheetProtectionPassword(state, 0, 'pw1')).toBe(true);
  });
});

describe('protect-sheet options as Excel measures them', () => {
  const setup = async (permissions: Record<string, boolean> = {}) => {
    const { WorkbookHandle } = await import('../../../src/engine/workbook-handle.js');
    const { mutators } = await import('../../../src/store/store.js');
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    for (let r = 0; r < 5; r += 1) {
      for (let c = 0; c < 3; c += 1)
        wb.setNumber({ sheet: 0, row: r, col: c }, (r + 1) * 10 + c + 1);
    }
    mutators.replaceCells(store, wb.cells(0));
    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });
    setProtectedSheet(store, 0, true, { permissions });
    return { wb, store };
  };
  const bold = async (store: ReturnType<typeof createSpreadsheetStore>) => {
    const { toggleBold } = await import('../../../src/commands/format.js');
    toggleBold(store.getState(), store);
    return store.getState().format.formats.get('0:2:0')?.bold === true;
  };
  const resize = async (store: ReturnType<typeof createSpreadsheetStore>) => {
    const { setColsWidth, setRowsHeight } = await import('../../../src/commands/row-col-layout.js');
    setColsWidth(store, null, 0, 0, 150);
    setRowsHeight(store, null, 0, 0, 40);
    const { colWidths, rowHeights } = store.getState().layout;
    return [colWidths.get(0) === 150, rowHeights.get(0) === 40];
  };

  it('blocks cell formatting by default and allows it with formatCells', async () => {
    expect(await bold((await setup()).store)).toBe(false);
    expect(await bold((await setup({ formatCells: true })).store)).toBe(true);
  });

  it('blocks column width and row height by default, each allowed by its option', async () => {
    expect(await resize((await setup()).store)).toEqual([false, false]);
    expect(await resize((await setup({ formatColumns: true })).store)).toEqual([true, false]);
    expect(await resize((await setup({ formatRows: true })).store)).toEqual([false, true]);
    // Formatting cells does not open widths or heights.
    expect(await resize((await setup({ formatCells: true })).store)).toEqual([false, false]);
  });

  it('allows inserting rows with insertRows', async () => {
    const { insertRows } = await import('../../../src/commands/structure.js');
    expect(insertRows((await setup()).store, (await setup()).wb, null, 1)).toBe(false);
    const allowed = await setup({ insertRows: true });
    expect(insertRows(allowed.store, allowed.wb, null, 1)).toBe(true);
    expect(allowed.wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({
      kind: 'number',
      value: 21,
    });
  });

  it('deletes rows with deleteRows only when every cell in them is unlocked', async () => {
    const { deleteRows } = await import('../../../src/commands/structure.js');
    const locked = await setup({ deleteRows: true });
    expect(deleteRows(locked.store, locked.wb, null, 4)).toBe(false);
    const unlocked = await setup({ deleteRows: true });
    setCellLocked(unlocked.store, { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 16383 }, false);
    expect(deleteRows(unlocked.store, unlocked.wb, null, 4)).toBe(true);
    const none = await setup();
    setCellLocked(none.store, { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 16383 }, false);
    expect(deleteRows(none.store, none.wb, null, 4)).toBe(false);
  });
});
