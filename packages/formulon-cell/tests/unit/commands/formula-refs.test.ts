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
    ['=XFE1+A1048577+XFD1048576', 1, 1, '=XFE1+A1048577+#REF!'],
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

  it('shifts whole-column and whole-row ranges on their relative axes', () => {
    expect(shiftFormulaRefs('=SUM(A:A,$B:$C,1:3,$4:$5)', 1, 1)).toBe('=SUM(B:B,$B:$C,2:4,$4:$5)');
    expect(shiftFormulaRefs('=SUM(XFD:XFD)', 0, 1)).toBe('=SUM(#REF!)');
    expect(shiftFormulaRefs('=SUM(XFE:XFE)', 0, 1)).toBe('=SUM(XFE:XFE)');
  });

  it('ignores refs inside string literals', () => {
    expect(shiftFormulaRefs('="A1"&A1', 1, 0)).toBe('="A1"&A2');
  });

  it('emits #REF! when a valid ref shifts outside the grid', () => {
    expect(shiftFormulaRefs('=A1', -5, 0)).toBe('=#REF!');
    expect(shiftFormulaRefs('=A1048576', 1, 0)).toBe('=#REF!');
    expect(shiftFormulaRefs('=XFD1', 0, 1)).toBe('=#REF!');
  });

  it('leaves originally invalid names and refs untouched', () => {
    expect(shiftFormulaRefs('=A1048577+Year2024', 1, 1)).toBe('=A1048577+Year2024');
  });

  it('leaves structured references untouched', () => {
    expect(shiftFormulaRefs('=T1[Q1]+T1[[#Headers],[Q1]]+$A$1', 1, 1)).toBe(
      '=T1[Q1]+T1[[#Headers],[Q1]]+$A$1',
    );
  });

  it('does not rewrite cell-looking text inside escaped table headers', () => {
    expect(shiftFormulaRefs("=T1[X']A1]+A1", 1, 0)).toBe("=T1[X']A1]+A2");
  });

  it('preserves Unicode and 3-D sheet qualifiers while shifting cells', () => {
    expect(shiftFormulaRefs("=シート1!A3+'シート 2'!A3=シート1:シート3!A3", 1, 0)).toBe(
      "=シート1!A4+'シート 2'!A4=シート1:シート3!A4",
    );
  });

  it('leaves external-workbook refs unsupported and unchanged', () => {
    expect(shiftFormulaRefs('=[Book.xlsx]Sheet1!A3', 1, 1)).toBe('=[Book.xlsx]Sheet1!A3');
    expect(shiftFormulaRefs("='[Book.xlsx]Sheet 1'!A3", 1, 1)).toBe("='[Book.xlsx]Sheet 1'!A3");
  });
});

describe('adjustFormulaForRowColEdit — insert/delete rows/cols', () => {
  it.each([
    ['=A3+B$3+$C3+$D$3', 'row', 2, 1, '=A4+B$4+$C4+$D$4'],
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

  it('shifts every mixed absolute row reference on row insertion', () => {
    expect(adjustFormulaForRowColEdit('=$A$3+A$3+$A3+A3', 'row', 2, 1)).toBe('=$A$4+A$4+$A4+A4');
  });

  it('turns every mixed absolute row reference into #REF! on row deletion', () => {
    expect(adjustFormulaForRowColEdit('=$A$4+A$4+$A4+A4', 'row', 3, -1)).toBe(
      '=#REF!+#REF!+#REF!+#REF!',
    );
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

  it('retains $ markers while structural edits move the edited axis', () => {
    expect(adjustFormulaForRowColEdit('=A$3', 'row', 0, 5)).toBe('=A$8');
    expect(adjustFormulaForRowColEdit('=$C1', 'col', 0, 5)).toBe('=$H1');
  });

  it('updates whole-axis ranges on matching structural edits regardless of $', () => {
    expect(adjustFormulaForRowColEdit('=SUM(A:A,$B:$C)', 'col', 0, 1)).toBe('=SUM(B:B,$C:$D)');
    expect(adjustFormulaForRowColEdit('=SUM(1:3,$4:$5)', 'row', 1, 1)).toBe('=SUM(1:4,$5:$6)');
    expect(adjustFormulaForRowColEdit('=SUM(1:3)', 'row', 1, -1)).toBe('=SUM(1:2)');
    expect(adjustFormulaForRowColEdit('=SUM(2:2)', 'row', 1, -1)).toBe('=SUM(#REF!)');
  });

  it('leaves whole-axis ranges unchanged for partial cell-band shifts', () => {
    const affected = { r0: 2, c0: 0, r1: 1048575, c1: 0 };
    expect(adjustFormulaForCellBandShift('=SUM(A:A,1:3)', affected, 'down', 1)).toBe(
      '=SUM(A:A,1:3)',
    );
  });

  it('leaves qualified Unicode refs untouched during structural edits', () => {
    expect(adjustFormulaForRowColEdit("=シート1!A3+'シート 2'!A3", 'row', 2, 1)).toBe(
      "=シート1!A3+'シート 2'!A3",
    );
  });

  it('adjusts cell-band refs by their resolved sheet binding', () => {
    const affected = { r0: 1, c0: 0, r1: 1048575, c1: 0 };
    const sourceFormula = {
      editedSheet: 0,
      formulaSheet: 0,
      sheetNames: ['Source', 'Target'],
    } as const;
    const targetFormula = { ...sourceFormula, formulaSheet: 1 };
    expect(
      adjustFormulaForCellBandShift(
        '=Source!$A$2+Source!A1:A3+A2',
        affected,
        'down',
        1,
        sourceFormula,
      ),
    ).toBe('=Source!$A$3+Source!A1:A4+A3');
    expect(
      adjustFormulaForCellBandShift(
        '=Source!$A$2+Source!A1:A3+A2',
        affected,
        'down',
        1,
        targetFormula,
      ),
    ).toBe('=Source!$A$3+Source!A1:A4+A2');
    expect(
      adjustFormulaForCellBandShift('=Source!A:A+Source!1:3', affected, 'down', 1, targetFormula),
    ).toBe('=Source!A:A+Source!1:3');
    expect(
      adjustFormulaForCellBandShift(
        '=Source!$A$3+Source!A1:A3',
        { ...affected, r0: 2 },
        'down',
        -1,
        targetFormula,
      ),
    ).toBe('=Source!#REF!+Source!A1:A2');
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

  it('shifts every mixed absolute row reference in a cell band', () => {
    const affected = { r0: 2, c0: 3, r1: 1048575, c1: 3 };
    expect(adjustFormulaForCellBandShift('=$D$3+D$3+$D3+D3', affected, 'down', 1)).toBe(
      '=$D$4+D$4+$D4+D4',
    );
  });

  it('shifts every mixed absolute column reference in a cell band', () => {
    const affected = { r0: 0, c0: 2, r1: 0, c1: 16383 };
    expect(adjustFormulaForCellBandShift('=$C$1+C$1+$C1+C1', affected, 'right', 1)).toBe(
      '=$D$1+D$1+$D1+D1',
    );
  });

  it('turns references inside a deleted cell band into #REF!', () => {
    const affected = { r0: 2, c0: 0, r1: 1048575, c1: 0 };
    expect(adjustFormulaForCellBandShift('=$A$3+A$3+$A3+A3', affected, 'up', -1)).toBe(
      '=#REF!+#REF!+#REF!+#REF!',
    );
    expect(adjustFormulaForCellBandShift('=A4', affected, 'up', -1)).toBe('=A3');
  });

  it('turns references inside a deleted cell band into #REF! horizontally', () => {
    const affected = { r0: 0, c0: 2, r1: 0, c1: 16383 };
    expect(adjustFormulaForCellBandShift('=$C$1+C$1+$C1+C1', affected, 'left', -1)).toBe(
      '=#REF!+#REF!+#REF!+#REF!',
    );
  });

  it('clamps a partially deleted range to the surviving boundary', () => {
    const affected = { r0: 2, c0: 0, r1: 1048575, c1: 0 };
    expect(adjustFormulaForCellBandShift('=SUM(A3:A20)', affected, 'up', -1)).toBe('=SUM(A3:A19)');
    expect(adjustFormulaForCellBandShift('=SUM(A1:A3)', affected, 'up', -1)).toBe('=SUM(A1:A2)');
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

  it('keeps partially overlapping same-sheet ranges intact', () => {
    const single = { r0: 0, c0: 0, r1: 0, c1: 0 };
    expect(adjustFormulaForCutPasteMove('=SUM(A1:A3)', single, { r0: 0, c0: 3 })).toBe(
      '=SUM(A1:A3)',
    );

    const middle = { r0: 1, c0: 0, r1: 1, c1: 0 };
    expect(adjustFormulaForCutPasteMove('=SUM(A1:A3)', middle, { r0: 1, c0: 8 })).toBe(
      '=SUM(A1:A3)',
    );
  });

  it('leaves string literals, function names, out-of-range refs, and sheet-qualified refs alone', () => {
    expect(adjustFormulaForCutPasteMove('="A1"&LOG10(A1)+C3+Sheet2!A1', source, dest)).toBe(
      '="A1"&LOG10(D5)+C3+Sheet2!A1',
    );
  });

  it('keeps moved formulas bound to their original sheet and follows external refs', () => {
    const source = { r0: 0, c0: 1, r1: 0, c1: 1 }; // Source!B1 → Target!D3
    const dest = { r0: 2, c0: 3 };
    const moved = {
      sourceSheet: 0,
      destinationSheet: 1,
      formulaSheet: 0,
      outputSheet: 1,
      sheetNames: ['Source', 'Target', 'Other'],
    } as const;
    expect(adjustFormulaForCutPasteMove('=A1', source, dest, moved)).toBe('=Source!A1');
    expect(adjustFormulaForCutPasteMove('=$A$1', source, dest, moved)).toBe('=Source!$A$1');
    expect(adjustFormulaForCutPasteMove('=A1:A2', source, dest, moved)).toBe('=Source!A1:A2');
    expect(adjustFormulaForCutPasteMove('=SUM(A:A,1:1)', source, dest, moved)).toBe(
      '=SUM(Source!A:A,Source!1:1)',
    );

    const external = { ...moved, formulaSheet: 0, outputSheet: 0 };
    expect(adjustFormulaForCutPasteMove('=B1+A1+Other!A1', source, dest, external)).toBe(
      '=Target!D3+A1+Other!A1',
    );
  });

  it('resolves qualified refs case-insensitively and quotes safe destination names', () => {
    const source = { r0: 0, c0: 1, r1: 0, c1: 1 };
    const dest = { r0: 2, c0: 3 };
    const context = {
      sourceSheet: 0,
      destinationSheet: 1,
      formulaSheet: 2,
      outputSheet: 2,
      sheetNames: ['Source', 'Target Sheet', 'Other', '対象'],
    } as const;
    expect(adjustFormulaForCutPasteMove("='sOuRcE'!B1", source, dest, context)).toBe(
      "='Target Sheet'!D3",
    );
    expect(adjustFormulaForCutPasteMove('=Source!B1+対象!A1', source, dest, context)).toBe(
      "='Target Sheet'!D3+対象!A1",
    );
    expect(adjustFormulaForCutPasteMove('=B1', source, dest, context)).toBe('=B1');
  });

  it('moves complete ranges and trims only cross-sheet edge strips', () => {
    const sourceCell = { r0: 0, c0: 0, r1: 0, c1: 0 };
    const crossSheet = {
      sourceSheet: 0,
      destinationSheet: 1,
      formulaSheet: 0,
      outputSheet: 0,
      sheetNames: ['Source', 'Target'],
    } as const;
    expect(
      adjustFormulaForCutPasteMove('=SUM(A1:A3)', sourceCell, { r0: 2, c0: 3 }, crossSheet),
    ).toBe('=SUM(A2:A3)');
    expect(
      adjustFormulaForCutPasteMove('=SUM(A1:B3)', sourceCell, { r0: 2, c0: 3 }, crossSheet),
    ).toBe('=SUM(A1:B3)');

    const middle = { r0: 1, c0: 5, r1: 1, c1: 5 };
    expect(adjustFormulaForCutPasteMove('=SUM(F1:F3)', middle, { r0: 2, c0: 3 }, crossSheet)).toBe(
      '=SUM(F1:F3)',
    );

    const complete = { r0: 1, c0: 0, r1: 2, c1: 0 };
    expect(
      adjustFormulaForCutPasteMove('=SUM(A2:A3)', complete, { r0: 1, c0: 2 }, crossSheet),
    ).toBe('=SUM(Target!C2:C3)');
  });

  it('keeps range endpoint order and absolute markers while trimming', () => {
    const sourceCell = { r0: 0, c0: 0, r1: 0, c1: 0 };
    const context = {
      sourceSheet: 0,
      destinationSheet: 1,
      formulaSheet: 0,
      outputSheet: 0,
      sheetNames: ['Source', 'Target'],
    } as const;
    expect(
      adjustFormulaForCutPasteMove('=SUM($A$3:$A$1)', sourceCell, { r0: 2, c0: 3 }, context),
    ).toBe('=SUM($A$3:$A$2)');
  });

  it('preserves ranges bound to other sheets and trims every source edge', () => {
    const context = {
      sourceSheet: 0,
      destinationSheet: 1,
      formulaSheet: 0,
      outputSheet: 0,
      sheetNames: ['Source', 'Target', 'Other'],
    };
    expect(
      adjustFormulaForCutPasteMove(
        '=SUM(Other!A1:B2)',
        { r0: 0, c0: 0, r1: 1, c1: 1 },
        { r0: 4, c0: 3 },
        context,
      ),
    ).toBe('=SUM(Other!A1:B2)');
    expect(
      adjustFormulaForCutPasteMove(
        '=SUM(A1:A3)',
        { r0: 2, c0: 0, r1: 2, c1: 0 },
        { r0: 4, c0: 3 },
        context,
      ),
    ).toBe('=SUM(A1:A2)');
    expect(
      adjustFormulaForCutPasteMove(
        '=SUM(A1:C2)',
        { r0: 0, c0: 0, r1: 1, c1: 0 },
        { r0: 4, c0: 3 },
        context,
      ),
    ).toBe('=SUM(B1:C2)');
    expect(
      adjustFormulaForCutPasteMove(
        '=SUM(A1:C2)',
        { r0: 0, c0: 2, r1: 1, c1: 2 },
        { r0: 4, c0: 3 },
        context,
      ),
    ).toBe('=SUM(A1:B2)');
  });

  it('preserves non-moved 3-D and invalid qualified refs', () => {
    const source = { r0: 0, c0: 1, r1: 0, c1: 1 };
    const dest = { r0: 2, c0: 3 };
    const context = {
      sourceSheet: 0,
      destinationSheet: 1,
      formulaSheet: 0,
      outputSheet: 1,
      sheetNames: ['Source', 'Target'],
    } as const;
    expect(
      adjustFormulaForCutPasteMove('=Source:Target!B1+Missing!B1', source, dest, context),
    ).toBe('=Source:Target!B1+Missing!B1');
  });
});
