import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { fillRange } from '../../../src/commands/fill.js';
import { shiftFormulaRefs } from '../../../src/commands/formula-refs.js';
import { addrKey, MAX_COL, MAX_ROW } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('fill of formulas against the desktop app (measured)', () => {
  let wb: WorkbookHandle;
  beforeAll(async () => {
    wb = await WorkbookHandle.createDefault();
  });
  afterAll(() => wb.dispose());

  const fill = (
    sheet: number,
    src: { row: number; col: number },
    dest: { row: number; col: number },
    formula: string,
  ): ((row: number, col: number) => string | null | undefined) => {
    const store = createSpreadsheetStore();
    const addr = { sheet, ...src };
    wb.setFormula(addr, formula);
    wb.recalc();
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set(addrKey(addr), { value: wb.getValue(addr), formula });
      return { ...s, data: { ...s.data, cells } };
    });
    const r0 = Math.min(src.row, dest.row);
    const r1 = Math.max(src.row, dest.row);
    const c0 = Math.min(src.col, dest.col);
    const c1 = Math.max(src.col, dest.col);
    fillRange(
      store.getState(),
      wb,
      { sheet, r0: src.row, c0: src.col, r1: src.row, c1: src.col },
      { sheet, r0, c0, r1, c1 },
    );
    return (row, col) =>
      [...wb.cells(sheet)].find((c) => c.addr.row === row && c.addr.col === col)?.formula;
  };

  it('keeps each $ with its coordinate when a filled range crosses (c20)', () => {
    const at = fill(0, { row: 70, col: 4 }, { row: 70, col: 0 }, '=SUM($G73:H$74)');
    expect(at(70, 0)).toBe('=SUM(D73:$G$74)');
    expect(at(70, 2)).toBe('=SUM(F73:$G$74)');
  });

  it('shiftFormulaRefs can follow the coordinate or keep the slot', () => {
    expect(shiftFormulaRefs('=SUM($G73:H$74)', 0, -4, { crossedRange: 'follow-coordinate' })).toBe(
      '=SUM(D73:$G$74)',
    );
    expect(shiftFormulaRefs('=SUM($G73:H$74)', 0, -4)).toBe('=SUM($D73:G$74)');
  });

  it('does not throw when a range reference runs off the sheet (c13)', () => {
    const at = fill(
      0,
      { row: 14, col: MAX_COL - 1 },
      { row: 14, col: MAX_COL },
      '=SUM($XFD17:XFD$18)',
    );
    expect(at(14, MAX_COL)).toBe('=SUM(#REF!)');
  });

  it('keeps the sheet qualifier on an off-sheet reference (c14)', () => {
    expect(shiftFormulaRefs('=Ref!E1048576', 1, 0)).toBe('=Ref!#REF!');
    expect(shiftFormulaRefs("='My Sheet'!A1048576", 1, 0)).toBe("='My Sheet'!#REF!");
    expect(shiftFormulaRefs('=A1048576', 1, 0)).toBe('=#REF!');
    expect(MAX_ROW).toBe(1048575);
  });
});
