import { describe, expect, it } from 'vitest';
import { suggestStructuredRef } from '../../../src/interact/autocomplete.js';

const sales = {
  name: 'Sales',
  columns: ['Region', 'Revenue', 'Quarter'],
};

describe('suggestStructuredRef', () => {
  it('suggests all columns just inside an empty bracket', () => {
    const text = '=Sales[';
    const ctx = suggestStructuredRef(text, text.length, [sales]);
    expect(ctx?.matches).toEqual(['Region', 'Revenue', 'Quarter']);
    expect(ctx?.tokenStart).toBe('=Sales['.length);
    expect(ctx?.tokenEnd).toBe('=Sales['.length);
    expect(ctx?.insertSuffix).toBe(']');
  });

  it('filters by partial prefix, case-insensitive', () => {
    const text = '=Sales[re';
    const ctx = suggestStructuredRef(text, text.length, [sales]);
    expect(ctx?.matches).toEqual(['Region', 'Revenue']);
  });

  it('returns null when caret is outside any open bracket', () => {
    expect(suggestStructuredRef('=SUM(A1)', 8, [sales])).toBeNull();
  });

  it('returns null when the table name does not match', () => {
    expect(suggestStructuredRef('=Other[r', 8, [sales])).toBeNull();
  });

  it('returns null when text is not a formula', () => {
    expect(suggestStructuredRef('Sales[r', 7, [sales])).toBeNull();
  });

  it('returns null when bracket already closed before caret', () => {
    expect(suggestStructuredRef('=Sales[Region]', 14, [sales])).toBeNull();
  });

  it('returns null when no tables are passed', () => {
    expect(suggestStructuredRef('=Sales[r', 8, [])).toBeNull();
  });
});
