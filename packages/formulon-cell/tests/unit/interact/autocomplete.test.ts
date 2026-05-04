import { describe, expect, it } from 'vitest';
import { suggestColumnHistory, suggestStructuredRef } from '../../../src/interact/autocomplete.js';

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

describe('suggestColumnHistory', () => {
  it('matches by case-insensitive prefix and replaces the whole token', () => {
    const text = 'ap';
    const ctx = suggestColumnHistory(text, text.length, ['Apple', 'Banana', 'Apricot']);
    expect(ctx?.matches).toEqual(['Apple', 'Apricot']);
    expect(ctx?.tokenStart).toBe(0);
    expect(ctx?.tokenEnd).toBe(text.length);
    expect(ctx?.insertSuffix).toBe('');
    expect(ctx?.kind).toBe('column');
  });

  it('preserves caller-supplied order (nearest-first) and dedupes', () => {
    // The values array is the editor's responsibility — it must come in
    // nearest-first / deduped. Verify the suggester keeps that order intact
    // and never re-emits a duplicate it sees later in the list.
    const ctx = suggestColumnHistory('a', 1, ['Apricot', 'Apple', 'Avocado']);
    expect(ctx?.matches).toEqual(['Apricot', 'Apple', 'Avocado']);
  });

  it('returns null when nothing prefix-matches', () => {
    expect(suggestColumnHistory('zz', 2, ['Apple', 'Banana'])).toBeNull();
  });

  it('returns null on empty input or mid-edit caret', () => {
    expect(suggestColumnHistory('', 0, ['Apple'])).toBeNull();
    // Caret at offset 1 of "apple" — user is correcting the middle, not
    // appending; Excel doesn't pop the list there either.
    expect(suggestColumnHistory('apple', 1, ['Apricot'])).toBeNull();
  });

  it('skips exact-length matches (would be a no-op insert)', () => {
    // "App" already typed — no point offering "App" itself, but longer
    // prefix-matches should still surface.
    const ctx = suggestColumnHistory('App', 3, ['App', 'Apple', 'Application']);
    expect(ctx?.matches).toEqual(['Apple', 'Application']);
  });
});
