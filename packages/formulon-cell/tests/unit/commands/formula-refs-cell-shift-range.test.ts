import { describe, expect, it } from 'vitest';
import { adjustFormulaForCellBandShift } from '../../../src/commands/formula-refs';

const END = 1048575;
const EXTENT = 16383;

describe('adjustFormulaForCellBandShift range coverage', () => {
  it('keeps a range the shifted band does not cover across the shift axis', () => {
    // insert shift right at C3 (single cell)
    const c3 = { r0: 2, c0: 2, r1: 2, c1: EXTENT };
    expect(adjustFormulaForCellBandShift('=SUM(C3:C8)', c3, 'right', 1)).toBe('=SUM(C3:C8)');
    expect(adjustFormulaForCellBandShift('=SUM(C3:E6)', c3, 'right', 1)).toBe('=SUM(C3:E6)');
    // insert shift right C3:D4
    const block = { r0: 2, c0: 2, r1: 3, c1: EXTENT };
    expect(adjustFormulaForCellBandShift('=SUM($C$3:$E$5)', block, 'right', 2)).toBe(
      '=SUM($C$3:$E$5)',
    );
    // insert shift down at E5
    const e5 = { r0: 4, c0: 4, r1: END, c1: 4 };
    expect(adjustFormulaForCellBandShift('=SUM($C$3:$E$5)', e5, 'down', 1)).toBe('=SUM($C$3:$E$5)');
    // insert shift down at D6
    const d6 = { r0: 5, c0: 3, r1: END, c1: 3 };
    expect(adjustFormulaForCellBandShift('=SUM(B7:D7)', d6, 'down', 1)).toBe('=SUM(B7:D7)');
    // insert shift right at D5
    const d5 = { r0: 4, c0: 3, r1: 4, c1: EXTENT };
    expect(adjustFormulaForCellBandShift('=SUM(C4:D5)', d5, 'right', 1)).toBe('=SUM(C4:D5)');
  });

  it('keeps ranges on delete shift up / left when the band is narrower than the range', () => {
    const d5Up = { r0: 4, c0: 3, r1: END, c1: 3 };
    expect(adjustFormulaForCellBandShift('=SUM(B7:D7)', d5Up, 'up', -1)).toBe('=SUM(B7:D7)');
    expect(adjustFormulaForCellBandShift('=SUM(C2:D5)', d5Up, 'up', -1)).toBe('=SUM(C2:D5)');
    const blockLeft = { r0: 4, c0: 3, r1: 5, c1: EXTENT };
    expect(adjustFormulaForCellBandShift('=SUM(C3:E6)', blockLeft, 'left', -2)).toBe('=SUM(C3:E6)');
  });

  it('still resizes a range the band fully covers across the shift axis', () => {
    const band = { r0: 2, c0: 2, r1: 7, c1: EXTENT };
    expect(adjustFormulaForCellBandShift('=SUM(C3:C8)', band, 'right', 1)).toBe('=SUM(D3:D8)');
  });
});
