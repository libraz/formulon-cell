import { describe, expect, it } from 'vitest';
import { applySubstitution, type FindOptions } from '../../../src/commands/find.js';

const options = (query: string, matchWidth?: boolean): FindOptions => ({
  query,
  ...(matchWidth === undefined ? {} : { matchWidth }),
});

describe('find: half-width and full-width forms', () => {
  it('treats full-width and half-width ASCII as equal when width is ignored', () => {
    expect(applySubstitution('ＡＢＣ１２３', options('abc123', false), 'x')).toBe('x');
    expect(applySubstitution('abc123', options('ＡＢＣ１２３', false), 'x')).toBe(
      'abc123'.replace(/.*/, 'x'),
    );
  });

  it('treats half-width and full-width katakana as equal when width is ignored', () => {
    expect(applySubstitution('ｱｲｳ', options('アイウ', false), '-')).toBe('-');
    expect(applySubstitution('アイウ', options('ｱｲｳ', false), '-')).toBe('-');
  });

  it('keeps the unmatched text untouched and the comparison exact by default', () => {
    expect(applySubstitution('ＡＢ-12', options('ab', false), '_')).toBe('_-12');
    expect(applySubstitution('ＡＢ-12', options('ab'), '_')).toBe('ＡＢ-12');
    expect(applySubstitution('ｱｲｳ', options('アイウ', true), '-')).toBe('ｱｲｳ');
  });
});
