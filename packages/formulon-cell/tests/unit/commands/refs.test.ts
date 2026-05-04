import { describe, expect, it } from 'vitest';
import {
  findActiveSignature,
  shiftFormulaRefs,
  suggestFunctions,
} from '../../../src/commands/refs.js';

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

describe('findActiveSignature', () => {
  it('returns null when caret is outside a function call', () => {
    expect(findActiveSignature('=A1+B2', 5)).toBeNull();
  });

  it('returns null when text is not a formula', () => {
    expect(findActiveSignature('SUM(1, 2)', 5)).toBeNull();
  });

  it('returns the matching signature for a known function', () => {
    const sig = findActiveSignature('=SUM(', 5);
    expect(sig?.name).toBe('SUM');
    expect(sig?.activeArgIndex).toBe(0);
  });

  it('bumps activeArgIndex once per top-level comma', () => {
    const text = '=IF(A1>0, B1,';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('IF');
    expect(sig?.activeArgIndex).toBe(2);
  });

  it('ignores commas inside nested calls', () => {
    const text = '=SUMIF(A1:A10, ">5", SUM(B1, B2),';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('SUMIF');
    expect(sig?.activeArgIndex).toBe(3);
  });

  it('ignores commas inside string literals', () => {
    const text = '=CONCAT("a, b", ';
    const sig = findActiveSignature(text, text.length);
    expect(sig?.name).toBe('CONCAT');
    expect(sig?.activeArgIndex).toBe(1);
  });

  it('returns null for unknown function names', () => {
    expect(findActiveSignature('=NOTAFUNC(', 10)).toBeNull();
  });
});

describe('suggestFunctions', () => {
  it('returns built-in matches by default', () => {
    const r = suggestFunctions('=SU', 3);
    expect(r?.token).toBe('SU');
    expect(r?.matches).toContain('SUM');
  });

  it('returns null when not in a formula', () => {
    expect(suggestFunctions('SUM', 3)).toBeNull();
  });

  it('uses opts.names when supplied (engine catalog override)', () => {
    const engineNames = ['CUSTOM_FN', 'CUSTOM_OTHER', 'OTHER'];
    const r = suggestFunctions('=CUS', 4, 8, { names: engineNames });
    expect(r?.matches).toEqual(['CUSTOM_FN', 'CUSTOM_OTHER']);
  });

  it('does not fall back to built-ins when opts.names is empty', () => {
    expect(suggestFunctions('=SU', 3, 8, { names: [] })).toBeNull();
  });
});
