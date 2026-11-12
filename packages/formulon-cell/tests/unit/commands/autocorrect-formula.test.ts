import { describe, expect, it } from 'vitest';
import { AutoCorrectStore } from '../../../src/commands/autocorrect.js';

const store = (): AutoCorrectStore =>
  new AutoCorrectStore(
    () => ({ entries: { en: [{ from: 'teh', to: 'the' }] } }),
    () => 'en',
    () => {},
  );

describe('AutoCorrect in formulas', () => {
  it('replaces a whole-word entry outside the strings, as the desktop does', () => {
    expect(store().correctEntry('=teh+1')).toBe('=the+1');
    expect(store().correctEntry('="teh"&teh')).toBe('="the"&the');
  });

  it('leaves names that only contain an entry, and the formula when none matches', () => {
    expect(store().correctEntry('=tehx+xteh')).toBe('=tehx+xteh');
    expect(store().correctEntry('=SUM(A1:B2)')).toBe('=SUM(A1:B2)');
  });

  it('applies only the list outside the strings, not the capitalisation rules', () => {
    expect(store().correctEntry('=sum(a1)')).toBe('=sum(a1)');
  });

  it('corrects outside the strings while typing', () => {
    expect(store().live('=teh+', 5)).toEqual({ value: '=the+', caret: 5 });
  });
});
