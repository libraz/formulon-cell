import { afterEach, describe, expect, it } from 'vitest';
import { isCellWritable, setProtectedSheet } from '../../../src/commands/protection.js';
import { deleteCols, deleteRows } from '../../../src/commands/structure.js';
import { hydrateAxisLocksFromEngine } from '../../../src/engine/protection-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/** Row and column styles decide the lock of a cell with no format of its own (Windows profile:
 *  no Mac chrome registered, so a delete needs every cell of the band unlocked). */
describe('locks inherited from row and column styles', () => {
  const handles: WorkbookHandle[] = [];

  afterEach(() => {
    for (const handle of handles.splice(0)) handle.dispose();
  });

  const setup = async () => {
    const wb = await WorkbookHandle.createDefault();
    handles.push(wb);
    for (let row = 0; row < 5; row += 1) wb.setNumber({ sheet: 0, row, col: 0 }, row + 1);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    const base = wb.getCellXf(0);
    if (!base) throw new Error('missing default XF');
    const unlocked = wb.addXfRecord({ ...base, locked: false });
    expect(unlocked).toBeGreaterThan(0);
    const sync = () => hydrateAxisLocksFromEngine(wb, store, 0);
    const value = (row: number, col = 0) => wb.getValue({ sheet: 0, row, col });
    return { wb, store, unlocked, sync, value };
  };

  it('deletes a row its row style unlocks, with Delete rows on', async () => {
    const { wb, store, unlocked, sync, value } = await setup();
    expect(wb.setRowStyleRange(0, 1, 1, unlocked)).toBe(true);
    sync();
    setProtectedSheet(store, 0, true, { permissions: { deleteRows: true } });
    expect(deleteRows(store, wb, null, 1)).toBe(true);
    expect(value(1)).toEqual({ kind: 'number', value: 3 });
  });

  it('keeps a row holding one locked cell', async () => {
    const { wb, store, unlocked, sync, value } = await setup();
    expect(wb.setRowStyleRange(0, 1, 1, unlocked)).toBe(true);
    sync();
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 5 }, { locked: true });
    setProtectedSheet(store, 0, true, { permissions: { deleteRows: true } });
    expect(deleteRows(store, wb, null, 1)).toBe(false);
    expect(value(1)).toEqual({ kind: 'number', value: 2 });
  });

  it('deletes a column its column style unlocks, and keeps one holding a locked cell', async () => {
    const { wb, store, unlocked, sync, value } = await setup();
    wb.setNumber({ sheet: 0, row: 0, col: 2 }, 30);
    wb.setNumber({ sheet: 0, row: 0, col: 3 }, 40);
    mutators.replaceCells(store, wb.cells(0));
    expect(wb.setColumnStyle(0, 1, 2, unlocked)).toBe(true);
    sync();
    mutators.setCellFormat(store, { sheet: 0, row: 4, col: 2 }, { locked: true });
    setProtectedSheet(store, 0, true, { permissions: { deleteColumns: true } });
    expect(deleteCols(store, wb, null, 2)).toBe(false);
    expect(value(0, 2)).toEqual({ kind: 'number', value: 30 });
    expect(deleteCols(store, wb, null, 1)).toBe(true);
    expect(value(0, 2)).toEqual({ kind: 'number', value: 40 });
  });

  it('lets a cell with no format of its own take its row style for a write', async () => {
    const { wb, store, unlocked, sync } = await setup();
    expect(wb.setRowStyleRange(0, 2, 2, unlocked)).toBe(true);
    sync();
    setProtectedSheet(store, 0, true);
    expect(isCellWritable(store.getState(), { sheet: 0, row: 2, col: 9 })).toBe(true);
    expect(isCellWritable(store.getState(), { sheet: 0, row: 3, col: 9 })).toBe(false);
  });
});
