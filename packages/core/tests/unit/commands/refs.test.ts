import { describe, expect, it } from 'vitest';
import { shiftFormulaRefs } from '../../../src/commands/refs.js';

describe('shiftFormulaRefs', () => {
  it('returns the input verbatim when delta is zero', () => {
    expect(shiftFormulaRefs('=A1+B2', 0, 0)).toBe('=A1+B2');
  });

  it('returns the input verbatim when not a formula', () => {
    expect(shiftFormulaRefs('A1+B2', 1, 1)).toBe('A1+B2');
  });

  it('shifts a single relative ref', () => {
    expect(shiftFormulaRefs('=A1', 2, 3)).toBe('=D3');
  });

  it('shifts a sum of relative refs', () => {
    expect(shiftFormulaRefs('=A1+B2', 1, 1)).toBe('=B2+C3');
  });

  it('respects $-locked column ($A1 — column anchored)', () => {
    expect(shiftFormulaRefs('=$A1', 2, 3)).toBe('=$A3');
  });

  it('respects $-locked row (A$1 — row anchored)', () => {
    expect(shiftFormulaRefs('=A$1', 2, 3)).toBe('=D$1');
  });

  it('does not shift a fully absolute ref ($A$1)', () => {
    expect(shiftFormulaRefs('=$A$1', 5, 5)).toBe('=$A$1');
  });

  it('handles ranges atom-by-atom', () => {
    expect(shiftFormulaRefs('=SUM(A1:B5)', 1, 1)).toBe('=SUM(B2:C6)');
  });

  it('handles a range with mixed locks', () => {
    expect(shiftFormulaRefs('=SUM($A1:B$5)', 2, 3)).toBe('=SUM($A3:E$5)');
  });

  it('does not touch refs inside string literals', () => {
    expect(shiftFormulaRefs('="A1 is here"+B2', 1, 1)).toBe('="A1 is here"+C3');
  });

  it('does not treat a function call as a ref (SIN(A1) — only A1 shifts)', () => {
    expect(shiftFormulaRefs('=SIN(A1)', 1, 1)).toBe('=SIN(B2)');
  });

  it('shifts every ref in a SUM range plus an addend', () => {
    expect(shiftFormulaRefs('=SUM(A1:B2)+C3', 1, 1)).toBe('=SUM(B2:C3)+D4');
  });

  it('shifts multi-letter columns (AA1 → AB1 with +1 col)', () => {
    expect(shiftFormulaRefs('=AA1', 0, 1)).toBe('=AB1');
  });

  it('handles negative shifts', () => {
    expect(shiftFormulaRefs('=C3', -1, -1)).toBe('=B2');
  });

  it('leaves out-of-range refs verbatim (engine surfaces #REF!)', () => {
    // Shifting A1 by (-1, 0) would yield row=-1 → invalid. We keep the source.
    expect(shiftFormulaRefs('=A1', -1, 0)).toBe('=A1');
  });

  it('handles a sheet-qualified ref (Sheet1!A1)', () => {
    // Atom regex doesn't match the sheet token, so only the cell part shifts.
    expect(shiftFormulaRefs('=Sheet1!A1', 1, 1)).toBe('=Sheet1!B2');
  });
});
