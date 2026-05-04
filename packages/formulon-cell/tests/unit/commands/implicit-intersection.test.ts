import { describe, expect, it } from 'vitest';
import {
  caretInsideImplicitIntersection,
  findImplicitIntersections,
} from '../../../src/commands/refs.js';

describe('findImplicitIntersections', () => {
  it('returns [] for non-formula text', () => {
    expect(findImplicitIntersections('@SUM(A1)')).toEqual([]);
  });

  it('finds a single leading @ before a function call', () => {
    expect(findImplicitIntersections('=@SUM(A1:A10)')).toEqual([{ at: 1, index: 0 }]);
  });

  it('finds multiple operators in source order', () => {
    expect(findImplicitIntersections('=@A1 + @SUM(B1:B5)')).toEqual([
      { at: 1, index: 0 },
      { at: 7, index: 1 },
    ]);
  });

  it('skips @ inside a string literal', () => {
    expect(findImplicitIntersections('="@A1"+B2')).toEqual([]);
  });

  it('skips @ inside a structured-ref bracket (Table[@col])', () => {
    expect(findImplicitIntersections('=Table1[@col]')).toEqual([]);
  });

  it('rejects a @ followed by punctuation that cannot be a ref', () => {
    expect(findImplicitIntersections('=@+1')).toEqual([]);
  });

  it('accepts a @ before an open paren (parenthesised expression)', () => {
    const ops = findImplicitIntersections('=@(A1+B1)');
    expect(ops.length).toBe(1);
  });
});

describe('caretInsideImplicitIntersection', () => {
  it('is true while caret sits in the operand right after @', () => {
    const text = '=@SUM(A1:A10)';
    expect(caretInsideImplicitIntersection(text, 5)).toBe(true);
  });

  it('is false at the @ itself (caret precedes the operator)', () => {
    expect(caretInsideImplicitIntersection('=@SUM(A1)', 1)).toBe(false);
  });

  it('is false in a formula with no @ operator', () => {
    expect(caretInsideImplicitIntersection('=SUM(A1)', 5)).toBe(false);
  });
});
