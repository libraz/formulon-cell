import { describe, expect, it } from 'vitest';
import {
  adjustFormulaForCellBandShift,
  adjustFormulaForCutPasteMove,
  adjustFormulaForRowColEdit,
  shiftFormulaRefs,
} from '../../../src/commands/formula-refs.js';

describe('shiftFormulaRefs — relative offset (fill / paste)', () => {
  it.each([
    ['=A1+B1', 2, 3, '=D3+E3'],
    ['=$A1+A$1+$A$1', 2, 5, '=$A3+F$1+$A$1'],
    ['=SUM(A1:B2,LOG10(C3),ATAN2(D4,E5))', 1, 1, '=SUM(B2:C3,LOG10(D4),ATAN2(E5,F6))'],
    ['="A1"&A1&"Sheet2!B2"', 1, 1, '="A1"&B2&"Sheet2!B2"'],
    ['=Sheet2!A1+Data!B2', 1, 1, '=Sheet2!B2+Data!C3'],
    ["='My Sheet'!A1:A3", 2, 0, "='My Sheet'!A3:A5"],
    ['=XFE1+A1048577+XFD1048576', 1, 1, '=XFE1+A1048577+XFD1048576'],
  ])('matches golden relative shift %#', (input, dRow, dCol, expected) => {
    expect(shiftFormulaRefs(input, dRow, dCol)).toBe(expected);
  });

  it('shifts relative refs by row/col delta', () => {
    expect(shiftFormulaRefs('=A1+B1', 1, 0)).toBe('=A2+B2');
    expect(shiftFormulaRefs('=A1', 0, 1)).toBe('=B1');
    expect(shiftFormulaRefs('=A1+B1', 2, 3)).toBe('=D3+E3');
  });

  it('keeps $-anchored axes pinned', () => {
    expect(shiftFormulaRefs('=$A$1', 3, 3)).toBe('=$A$1');
    expect(shiftFormulaRefs('=$A1', 2, 5)).toBe('=$A3');
    expect(shiftFormulaRefs('=A$1', 2, 5)).toBe('=F$1');
  });

  it('does NOT mangle function names ending in digits (C-1)', () => {
    expect(shiftFormulaRefs('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)');
    expect(shiftFormulaRefs('=ATAN2(A1,B1)', 1, 0)).toBe('=ATAN2(A2,B2)');
    expect(shiftFormulaRefs('=LOG10(A1)', 0, 1)).toBe('=LOG10(B1)');
  });

  it('preserves sheet name while shifting the qualified ref (C-2 for fill)', () => {
    expect(shiftFormulaRefs('=Sheet2!A1', 1, 0)).toBe('=Sheet2!A2');
    expect(shiftFormulaRefs("='My Sheet'!A1", 1, 0)).toBe("='My Sheet'!A2");
    expect(shiftFormulaRefs('=Data!A5', 2, 0)).toBe('=Data!A7');
  });

  it('shifts both endpoints of a range', () => {
    expect(shiftFormulaRefs('=SUM(A1:A3)', 1, 0)).toBe('=SUM(A2:A4)');
    expect(shiftFormulaRefs('=SUM(A1:B2)', 0, 1)).toBe('=SUM(B1:C2)');
  });

  it('ignores refs inside string literals', () => {
    expect(shiftFormulaRefs('="A1"&A1', 1, 0)).toBe('="A1"&A2');
  });

  it('leaves out-of-grid shifts as the original text', () => {
    expect(shiftFormulaRefs('=A1', -5, 0)).toBe('=A1');
  });
});

describe('adjustFormulaForRowColEdit — insert/delete rows/cols', () => {
  it.each([
    ['=A3+B$3+$C3+$D$3', 'row', 2, 1, '=A4+B$3+$C4+$D$3'],
    ['=A3:B5', 'row', 3, -1, '=A3:B4'],
    ['=A3:B5', 'row', 2, -3, '=#REF!'],
    ['=B1:D1', 'col', 2, -1, '=B1:C1'],
    ['=LOG10(A3)+Year2024', 'row', 2, 1, '=LOG10(A4)+Year2024'],
    ['=Sheet2!A3+Local!B4+A3', 'row', 2, 1, '=Sheet2!A3+Local!B4+A4'],
    ["='My Sheet'!A3+A3", 'row', 2, 1, "='My Sheet'!A3+A4"],
  ])('matches golden structural edit %#', (input, axis, split, delta, expected) => {
    expect(adjustFormulaForRowColEdit(input, axis as 'row' | 'col', split, delta)).toBe(expected);
  });

  it('shifts refs at/after an inserted row', () => {
    // insert 1 row at index 2 (split=2, delta=+1)
    expect(adjustFormulaForRowColEdit('=A3', 'row', 2, 1)).toBe('=A4');
    expect(adjustFormulaForRowColEdit('=A1', 'row', 2, 1)).toBe('=A1');
  });

  it('shifts refs at/after an inserted column', () => {
    expect(adjustFormulaForRowColEdit('=C1', 'col', 1, 1)).toBe('=D1');
    expect(adjustFormulaForRowColEdit('=A1', 'col', 1, 1)).toBe('=A1');
  });

  it('turns a single ref inside a deleted band into #REF!', () => {
    // delete row index 2 (split=2, delta=-1) → A3 (row idx2) is deleted
    expect(adjustFormulaForRowColEdit('=A3', 'row', 2, -1)).toBe('=#REF!');
    // A4 (row idx3) shifts up to A3
    expect(adjustFormulaForRowColEdit('=A4', 'row', 2, -1)).toBe('=A3');
  });

  it('clamps a range whose top endpoint is deleted rather than #REF! mid-range (H-1)', () => {
    // =SUM(A5:A20): delete rows 4..5 (split=4 [A5], delta=-1 removes only idx4)
    // A5 (idx4) deleted → clamp to boundary; A20 (idx19) shifts to idx18 → A19
    expect(adjustFormulaForRowColEdit('=SUM(A5:A20)', 'row', 4, -1)).toBe('=SUM(A5:A19)');
  });

  it('#REF!s a range only when both endpoints are deleted', () => {
    // =SUM(A5:A6): delete rows idx4..idx5 (split=4, delta=-2) → Excel: =SUM(#REF!)
    expect(adjustFormulaForRowColEdit('=SUM(A5:A6)', 'row', 4, -2)).toBe('=SUM(#REF!)');
  });

  it('leaves cross-sheet refs untouched (C-2)', () => {
    expect(adjustFormulaForRowColEdit('=Sheet2!A1', 'row', 0, 1)).toBe('=Sheet2!A1');
    expect(adjustFormulaForRowColEdit("='My Sheet'!A5", 'row', 0, 1)).toBe("='My Sheet'!A5");
    expect(adjustFormulaForRowColEdit('=Data!A5*2', 'row', 0, 5)).toBe('=Data!A5*2');
  });

  it('does not mistake function names / name-like tokens for refs', () => {
    expect(adjustFormulaForRowColEdit('=LOG10(A3)', 'row', 2, 1)).toBe('=LOG10(A4)');
    expect(adjustFormulaForRowColEdit('=Year2024*2', 'row', 0, 1)).toBe('=Year2024*2');
  });

  it('keeps $-anchored refs on the pinned axis', () => {
    expect(adjustFormulaForRowColEdit('=A$3', 'row', 0, 5)).toBe('=A$3');
    expect(adjustFormulaForRowColEdit('=$C1', 'col', 0, 5)).toBe('=$C1');
  });
});

describe('adjustFormulaForCellBandShift — insert/delete cells', () => {
  it('shifts refs inside the band down', () => {
    const affected = { r0: 2, c0: 0, r1: 1048575, c1: 0 };
    expect(adjustFormulaForCellBandShift('=A3', affected, 'down', 1)).toBe('=A4');
    // outside the column band → untouched
    expect(adjustFormulaForCellBandShift('=B3', affected, 'down', 1)).toBe('=B3');
  });

  it('does not mangle sheet names on cell-band shift', () => {
    const affected = { r0: 2, c0: 0, r1: 1048575, c1: 0 };
    expect(adjustFormulaForCellBandShift('=Sheet2!A3', affected, 'down', 1)).toBe('=Sheet2!A3');
  });

  it('does not mangle function names ending in digits', () => {
    const affected = { r0: 0, c0: 0, r1: 1048575, c1: 0 };
    expect(adjustFormulaForCellBandShift('=LOG10(A3)', affected, 'down', 1)).toBe('=LOG10(A4)');
  });
});

describe('adjustFormulaForCutPasteMove — external refs follow moved cells', () => {
  const source = { r0: 0, c0: 0, r1: 1, c1: 1 };
  const dest = { r0: 4, c0: 3 };

  it('moves references inside the cut source range to the pasted destination', () => {
    expect(adjustFormulaForCutPasteMove('=A1+B2+C3', source, dest)).toBe('=D5+E6+C3');
  });

  it('moves absolute references because the referenced cell moved', () => {
    expect(adjustFormulaForCutPasteMove('=$A$1+A$2+$B1', source, dest)).toBe('=$D$5+D$6+$E5');
  });

  it('moves range endpoints that overlap the cut source range', () => {
    expect(adjustFormulaForCutPasteMove('=SUM(A1:B2)', source, dest)).toBe('=SUM(D5:E6)');
  });

  it('leaves string literals, function names, out-of-range refs, and sheet-qualified refs alone', () => {
    expect(adjustFormulaForCutPasteMove('="A1"&LOG10(A1)+C3+Sheet2!A1', source, dest)).toBe(
      '="A1"&LOG10(D5)+C3+Sheet2!A1',
    );
  });
});
