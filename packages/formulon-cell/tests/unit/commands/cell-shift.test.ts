import { describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { History, undo } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

describe('cell shift commands', () => {
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
});
