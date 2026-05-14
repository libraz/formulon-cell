import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { boundingRange, findMatchingCells } from '../../src/commands/goto-special.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

/**
 * Integration: Go To Special. The dialog feeds these predicates against the
 * live mount; this test exercises the predicate matrix against the stub
 * engine so the formula / constant / text / error / blank paths each surface
 * the right candidate set.
 */
describe('integration: go-to-special', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet();
    const { instance, workbook } = sheet;
    // Layout (sheet 0):
    //   A1=1 (number)    B1="foo" (text)     C1=#REF! (text sentinel)
    //   A2=2 (number)    B2=blank            C2="=A1+A2" (formula → 3)
    //   A3=3 (number)    B3="bar" (text)     C3=blank
    workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    workbook.setText({ sheet: 0, row: 0, col: 1 }, 'foo');
    workbook.setText({ sheet: 0, row: 0, col: 2 }, '#REF!');
    workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 2);
    workbook.setFormula({ sheet: 0, row: 1, col: 2 }, '=A1+A2');
    workbook.setNumber({ sheet: 0, row: 2, col: 0 }, 3);
    workbook.setText({ sheet: 0, row: 2, col: 1 }, 'bar');
    workbook.recalc();
    mutators.replaceCells(instance.store, workbook.cells(0));
  });

  afterEach(() => sheet.dispose());

  it('formulas: returns only cells holding a formula', () => {
    const { instance, workbook } = sheet;
    const hits = findMatchingCells(workbook, instance.store, 'sheet', 'formulas');
    expect(hits).toEqual([{ sheet: 0, row: 1, col: 2 }]);
  });

  it('constants: returns numbers + text but excludes the formula', () => {
    const { instance, workbook } = sheet;
    const hits = findMatchingCells(workbook, instance.store, 'sheet', 'constants');
    // 6 constants: A1, A2, A3 (numbers) + B1, B3 (text) + C1 (text sentinel
    // counts as a constant since it's a typed string, not a live formula).
    expect(hits.length).toBe(6);
    expect(hits.some((a) => a.row === 1 && a.col === 2)).toBe(false);
  });

  it('numbers: returns only number-valued cells', () => {
    const { instance, workbook } = sheet;
    const hits = findMatchingCells(workbook, instance.store, 'sheet', 'numbers');
    const cols = hits.map((h) => `${h.row}:${h.col}`).sort();
    // A1, A2, A3 — plus C2 (formula evaluates to a number = 3).
    expect(cols).toEqual(['0:0', '1:0', '1:2', '2:0']);
  });

  it('text: returns plain text cells but excludes error sentinels', () => {
    const { instance, workbook } = sheet;
    const hits = findMatchingCells(workbook, instance.store, 'sheet', 'text');
    // B1 + B3. C1 is "#REF!" — typed text, but classed as error.
    const keys = hits.map((h) => `${h.row}:${h.col}`).sort();
    expect(keys).toEqual(['0:1', '2:1']);
  });

  it('errors: matches typed error sentinels', () => {
    const { instance, workbook } = sheet;
    const hits = findMatchingCells(workbook, instance.store, 'sheet', 'errors');
    expect(hits).toEqual([{ sheet: 0, row: 0, col: 2 }]);
  });

  it('blanks scoped to selection: surfaces empty interior cells', () => {
    const { instance, workbook } = sheet;
    // Selection over the full 3x3 grid.
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
    const hits = findMatchingCells(workbook, instance.store, 'selection', 'blanks');
    // B2 and C3 are the two empty cells in the rectangle.
    const keys = hits.map((h) => `${h.row}:${h.col}`).sort();
    expect(keys).toEqual(['1:1', '2:2']);
  });

  it('boundingRange: encloses the matches', () => {
    const matches = [
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 2, col: 1 },
      { sheet: 0, row: 1, col: 2 },
    ];
    expect(boundingRange(matches)).toEqual({ sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 });
  });
});
