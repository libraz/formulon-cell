import { describe, expect, it } from 'vitest';
import { createFormulaReferenceGesture } from '../../../src/interact/formula-reference-gesture.js';

describe('formula reference gesture', () => {
  it('applies every drag update to the original selected endpoint', () => {
    const formula = '=SUM(B1:B2)+D1';
    const gesture = createFormulaReferenceGesture(formula, 8, 10, {});

    expect(gesture.insert('C4')).toEqual({ text: '=SUM(B1:C4)+D1', caret: 10 });
    expect(gesture.insert('C4:D5')).toEqual({ text: '=SUM(B1:C4:D5)+D1', caret: 13 });
    expect(gesture.insert('A4:C5')).toEqual({ text: '=SUM(B1:A4:C5)+D1', caret: 13 });
  });

  it('preserves the untouched endpoint when the left endpoint is selected', () => {
    const gesture = createFormulaReferenceGesture('=SUM(B1:B2)+D1', 5, 7, {});

    expect(gesture.insert('C4:D5').text).toBe('=SUM(C4:D5:B2)+D1');
  });

  it('keeps whole-range and collapsed-range replacement as ordinary references', () => {
    const formula = '=SUM(B1:B2)+D1';

    expect(createFormulaReferenceGesture(formula, 5, 10, {}).insert('C4:D5').text).toBe(
      '=SUM(C4:D5)+D1',
    );
    expect(createFormulaReferenceGesture(formula, 8, 8, {}).insert('C4:D5').text).toBe(
      '=SUM(C4:D5)+D1',
    );
  });

  it('retains the R1C1 prefix and suffix from the gesture snapshot', () => {
    const formula = '=SUM(R[-1]C[-1]:R[0]C[0])+R[3]C[4]';
    const endpoint = formula.indexOf('R[0]C[0]');
    const gesture = createFormulaReferenceGesture(formula, endpoint, endpoint + 'R[0]C[0]'.length, {
      r1c1: true,
    });

    expect(gesture.insert('R[1]C[1]').text).toBe('=SUM(R[-1]C[-1]:R[1]C[1])+R[3]C[4]');
    expect(gesture.insert('R[1]C[1]:R[2]C[2]').text).toBe(
      '=SUM(R[-1]C[-1]:R[1]C[1]:R[2]C[2])+R[3]C[4]',
    );
  });
});
