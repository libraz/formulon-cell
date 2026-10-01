import { describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { History, undo } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

describe('cell shift commands', () => {
  it('preserves static errors through insert, delete and undo/redo', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    const history = new History();
    const source = { sheet: 0, row: 1, col: 0 };
    const shifted = { ...source, row: 2 };
    const range = { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 };
    wb.setError(source, 3);

    expect(insertCells(store, wb, history, range, 'down')).toBe(true);
    expect(wb.getValue(shifted)).toEqual({ kind: 'error', code: 3, text: '#REF!' });
    history.undo();
    expect(wb.getValue(source)).toEqual({ kind: 'error', code: 3, text: '#REF!' });
    history.redo();
    expect(wb.getValue(shifted)).toEqual({ kind: 'error', code: 3, text: '#REF!' });

    expect(deleteCells(store, wb, history, range, 'up')).toBe(true);
    expect(wb.getValue(source)).toEqual({ kind: 'error', code: 3, text: '#REF!' });
    history.undo();
    expect(wb.getValue(shifted)).toEqual({ kind: 'error', code: 3, text: '#REF!' });
    wb.dispose();
  });

  it('inserts cells by shifting only selected columns down', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setText({ sheet: 0, row: 1, col: 1 }, 'B2');
    wb.setText({ sheet: 0, row: 1, col: 2 }, 'C2');
    wb.setText({ sheet: 0, row: 1, col: 3 }, 'D2');
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, { bold: true });

    insertCells(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 }, 'down');

    expect(wb.getValue({ sheet: 0, row: 1, col: 1 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 }).kind).toBe('blank');
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'B2' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'text', value: 'C2' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 3 })).toEqual({ kind: 'text', value: 'D2' });
    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 1 }))?.bold).toBe(
      true,
    );
  });

  it('deletes cells by shifting only selected rows left', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setText({ sheet: 0, row: 1, col: 1 }, 'B2');
    wb.setText({ sheet: 0, row: 1, col: 2 }, 'C2');
    wb.setText({ sheet: 0, row: 2, col: 1 }, 'B3');
    wb.setText({ sheet: 0, row: 2, col: 2 }, 'C3');

    deleteCells(store, wb, null, { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 1 }, 'left');

    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'C2' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'C3' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 }).kind).toBe('blank');
  });

  it('groups cell shift undo into one history entry', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    const history = new History();
    wb.setText({ sheet: 0, row: 1, col: 1 }, 'B2');
    wb.setText({ sheet: 0, row: 2, col: 1 }, 'B3');

    insertCells(store, wb, history, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, 'down');
    undo(history);

    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'B2' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'B3' });
  });

  it('updates qualified references on every sheet and restores them as one undo step', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    expect(wb.addSheet('Target')).toBe(1);
    const history = new History();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    wb.setNumber({ sheet: 0, row: 2, col: 0 }, 3);
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=Sheet1!$A$2');
    wb.setFormula({ sheet: 1, row: 0, col: 3 }, '=Sheet1!$A$2');
    wb.setFormula({ sheet: 1, row: 1, col: 3 }, '=SUM(Sheet1!A1:A3)');
    wb.setFormula({ sheet: 1, row: 2, col: 3 }, '=A2');

    const range = { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 };
    expect(insertCells(store, wb, history, range, 'down')).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A4)');
    expect(wb.cellFormula({ sheet: 1, row: 2, col: 3 })).toBe('=A2');

    expect(history.undo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$2');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!$A$2');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A3)');
    expect(history.redo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A4)');

    const deleteRange = { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 };
    expect(deleteCells(store, wb, history, deleteRange, 'up')).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 0, col: 3 })).toBe('=Sheet1!#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A3)');
    expect(wb.cellFormula({ sheet: 1, row: 2, col: 3 })).toBe('=A2');
    expect(history.undo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!$A$3');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A4)');
    expect(history.redo()).toBe(true);
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 2 })).toBe('=Sheet1!#REF!');
    expect(wb.cellFormula({ sheet: 1, row: 1, col: 3 })).toBe('=SUM(Sheet1!A1:A3)');
    wb.dispose();
  });

  it('rejects an insert that would drop a last-row cell or format atomically', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    const history = new History();
    const last = { sheet: 0, row: 1048575, col: 0 };
    wb.setNumber(last, 7);
    mutators.setCellFormat(store, last, { bold: true, comment: 'edge' });

    expect(insertCells(store, wb, history, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 }, 'down')).toBe(
      false,
    );
    expect(wb.getValue(last)).toEqual({ kind: 'number', value: 7 });
    expect(store.getState().format.formats.get('0:1048575:0')).toMatchObject({
      bold: true,
      comment: 'edge',
    });
    expect(history.canUndo()).toBe(false);
    wb.dispose();
  });
});
