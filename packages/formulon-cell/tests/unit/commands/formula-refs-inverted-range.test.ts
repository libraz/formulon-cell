import { describe, expect, it } from 'vitest';
import { shiftFormulaRefs } from '../../../src/commands/formula-refs';

describe('shiftFormulaRefs inverted ranges', () => {
  it('swaps reversed rows and keeps each $ in its slot', () => {
    expect(shiftFormulaRefs('=SUM(A$33:A36)', -4, 0)).toBe('=SUM(A$32:A33)');
  });

  it('swaps reversed columns and keeps each $ in its slot', () => {
    expect(shiftFormulaRefs('=SUM($B74:E74)', 0, -4)).toBe('=SUM($A74:B74)');
  });

  it('swaps both axes independently', () => {
    expect(shiftFormulaRefs('=SUM(C3:$A$1)', 4, 4)).toBe('=SUM(A1:$G$7)');
  });

  it('leaves ranges that keep their order unchanged', () => {
    expect(shiftFormulaRefs('=SUM(A$3:A5)', 2, 1)).toBe('=SUM(B$3:B7)');
  });
});
