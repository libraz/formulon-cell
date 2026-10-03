import { describe, expect, it } from 'vitest';
import {
  __testing,
  deleteCols,
  deleteRows,
  insertCols,
  insertRows,
} from '../../../../src/commands/structure.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { newWb } from './fixtures.js';

describe('shiftFormulaRefs', () => {
  const shift = __testing.shiftFormulaRefs;

  it('shifts plain row refs forward on insert', () => {
    expect(shift('=A1+B5', 'row', 0, 1)).toBe('=A2+B6');
    expect(shift('=A1+B5', 'row', 3, 1)).toBe('=A1+B6'); // A1 untouched
  });

  it('shifts plain col refs forward on insert', () => {
    expect(shift('=A1+B1', 'col', 0, 1)).toBe('=B1+C1');
    expect(shift('=A1+C1', 'col', 1, 1)).toBe('=A1+D1');
  });

  it('shifts row and column refs, including absolute axes', () => {
    expect(shift('=$A$1+B5', 'row', 0, 1)).toBe('=$A$2+B6');
    expect(shift('=A$1+B5', 'row', 0, 1)).toBe('=A$2+B6');
    expect(shift('=$A1+$B5', 'row', 0, 1)).toBe('=$A2+$B6');
    expect(shift('=$A$1+$B$5', 'col', 0, 2)).toBe('=$C$1+$D$5');
  });

  it('handles ranges', () => {
    expect(shift('=SUM(A1:A10)', 'row', 0, 2)).toBe('=SUM(A3:A12)');
    expect(shift('=SUM($A$1:$A$10)', 'row', 0, 2)).toBe('=SUM($A$3:$A$12)');
  });

  it('produces #REF! for refs in deleted band', () => {
    // Delete row 2 (atRow=2, n=1, delta=-1). Refs to row 2 (0-indexed: 1) → #REF!
    expect(shift('=A2', 'row', 1, -1)).toBe('=#REF!');
    expect(shift('=A1+A2+A3', 'row', 1, -1)).toBe('=A1+#REF!+A2');
  });

  it('shifts AA, ZZ, etc.', () => {
    expect(shift('=AA1', 'col', 0, 1)).toBe('=AB1');
    expect(shift('=Z1', 'col', 0, 1)).toBe('=AA1');
  });

  it('leaves string literals untouched', () => {
    expect(shift('="A1"&B5', 'row', 0, 1)).toBe('="A1"&B6');
    expect(shift('="say ""hi"" A1"', 'row', 0, 1)).toBe('="say ""hi"" A1"');
  });

  it('returns the input unchanged when delta is 0', () => {
    expect(shift('=A1+B2', 'row', 0, 0)).toBe('=A1+B2');
  });

  it('preserves function names that look like prefixes', () => {
    // SUM is followed by `(` — should not be misread as a ref.
    expect(shift('=SUM(A1:A5)+IF(B1>0,1,0)', 'row', 0, 1)).toBe('=SUM(A2:A6)+IF(B2>0,1,0)');
  });

  it('handles lowercase letters by upper-casing the output label', () => {
    expect(shift('=a1', 'row', 0, 1)).toBe('=A2');
  });
});

describe('insertRows: formula ref shifting', () => {
  it('rewrites refs in cells that move down', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 5);
    wb.setFormula({ sheet: 0, row: 1, col: 0 }, '=A1*2');
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 10 });

    insertRows(store, wb, null, 0, 1); // insert above row 0
    // Formula moved from row 1 → row 2; ref A1 → A2
    expect(wb.cellFormula({ sheet: 0, row: 2, col: 0 })).toBe('=A2*2');
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'number', value: 10 });
  });

  it('rewrites refs in stationary cells that point past the insert split', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    // Stationary formula in row 0 referencing row 5.
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=A6');
    wb.setNumber({ sheet: 0, row: 5, col: 0 }, 99);
    wb.recalc();

    insertRows(store, wb, null, 3, 2); // insert 2 rows at row 3
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=A8');
    expect(wb.getValue({ sheet: 0, row: 7, col: 0 })).toEqual({ kind: 'number', value: 99 });
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 99 });
  });
});

describe('deleteRows: formula ref shifting and #REF!', () => {
  it('replaces refs to deleted rows with #REF!', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=A2');
    wb.setNumber({ sheet: 0, row: 1, col: 0 }, 7);
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 7 });

    deleteRows(store, wb, null, 1, 1); // delete row 1
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=#REF!');
  });

  it('shifts refs that point past the deletion band', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=A6');
    wb.setNumber({ sheet: 0, row: 5, col: 0 }, 42);
    wb.recalc();

    deleteRows(store, wb, null, 1, 2); // delete rows 1, 2
    // Stationary formula at row 0; A6 (row=5) → A4 (row=3)
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=A4');
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'number', value: 42 });
  });
});

describe('cross-sheet refs are preserved across structural edits (C-2)', () => {
  it('leaves a Sheet2!-qualified ref untouched while shifting local refs', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    // Stationary formula on sheet 0 referencing both the current sheet and Sheet2.
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=A6+Sheet2!A6');
    wb.recalc();

    insertRows(store, wb, null, 3, 2); // insert 2 rows at row 3 on the active sheet
    // Local A6 shifts to A8; the cross-sheet Sheet2!A6 must NOT move.
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=A8+Sheet2!A6');
  });

  it('does not corrupt the sheet name on column inserts', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=Sheet2!B2');
    wb.recalc();
    insertCols(store, wb, null, 0, 1);
    // Formula moved from col 0 → col 1; the Sheet2 ref stays verbatim.
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=Sheet2!B2');
  });
});

describe('insertCols / deleteCols: formula ref shifting', () => {
  it('insertCols shifts col refs', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 4);
    wb.setFormula({ sheet: 0, row: 0, col: 2 }, '=A1*5');
    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 })).toEqual({ kind: 'number', value: 20 });

    insertCols(store, wb, null, 1, 1);
    // Cell moved from col 2 → col 3; A1 unaffected (col=0 < split=1)
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 3 })).toBe('=A1*5');
    expect(wb.getValue({ sheet: 0, row: 0, col: 3 })).toEqual({ kind: 'number', value: 20 });
  });

  it('deleteCols replaces refs to deleted cols with #REF!', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setFormula({ sheet: 0, row: 0, col: 0 }, '=B1');
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 3);
    wb.recalc();

    deleteCols(store, wb, null, 1, 1); // delete col 1
    expect(wb.cellFormula({ sheet: 0, row: 0, col: 0 })).toBe('=#REF!');
  });
});
