import { describe, expect, it, vi } from 'vitest';
import {
  type AutoCorrectOptions,
  AutoCorrectStore,
  autoCorrectBefore,
  autoCorrectText,
  builtInAutoCorrectEntries,
} from '../../../src/commands/autocorrect.js';

const symbols = builtInAutoCorrectEntries('en');

describe('AutoCorrect', () => {
  it('starts from the desktop symbol replacements, and the two Japanese greetings', () => {
    expect(symbols.map((entry) => `${entry.from}→${entry.to}`)).toEqual([
      '(c)→©',
      '(e)→€',
      '(r)→®',
      '(tm)→™',
      '...→…',
    ]);
    expect(builtInAutoCorrectEntries('de')).toEqual(symbols);
    expect(builtInAutoCorrectEntries('fr')).toEqual(symbols);
    expect(builtInAutoCorrectEntries('zh').map((entry) => entry.from)).toEqual([
      '(c)',
      '(r)',
      '(tm)',
      '...',
    ]);
    expect(builtInAutoCorrectEntries('ja')).toEqual([
      { from: 'こんにちわ', to: 'こんにちは' },
      { from: 'こんばんわ', to: 'こんばんは' },
    ]);
    expect(builtInAutoCorrectEntries('ko')).toEqual([]);
    expect(builtInAutoCorrectEntries('th')).toEqual([]);
  });

  // The matching rule is not measured against the desktop yet.
  it('replaces an entry standing as a word, at the start or after white space', () => {
    expect(autoCorrectText('(c) 2026 Acme(tm)', symbols)).toBe('© 2026 Acme(tm)');
    expect(autoCorrectText('Wait ... (r).', symbols)).toBe('Wait … ®.');
    expect(autoCorrectText('Wait...', symbols)).toBe('Wait...');
    expect(autoCorrectText('(c)x', symbols)).toBe('(c)x');
  });

  it('matches a lower-case entry in any case and carries the capitals over', () => {
    const entries = [{ from: 'teh', to: 'the' }];
    expect(autoCorrectText('teh Teh TEH', entries)).toBe('the The THE');
    expect(autoCorrectText('ACME acme', [{ from: 'ACME', to: 'Acme Corp' }])).toBe(
      'Acme Corp acme',
    );
  });

  it('finds the entry typed just before a delimiter, as replacing while typing does', () => {
    expect(autoCorrectBefore('see (c) ', 7, symbols)).toEqual({ start: 4, end: 7, to: '©' });
    expect(autoCorrectBefore('see (c)x', 7, symbols)).toBeNull();
    expect(autoCorrectBefore('a(c) ', 4, symbols)).toBeNull();
  });

  it('keeps edits for the session and hands the whole state to the host', () => {
    const onChange = vi.fn();
    let options: AutoCorrectOptions = { onChange };
    let locale = 'en';
    const store = new AutoCorrectStore(
      () => options,
      () => locale,
    );
    expect(store.state()).toEqual({
      entries: symbols,
      replaceAsYouType: true,
      replaceHyperlinks: true,
    });
    store.add('teh', 'the');
    expect(store.correct('teh (c)')).toBe('the ©');
    expect(onChange).toHaveBeenLastCalledWith(
      {
        entries: [...symbols, { from: 'teh', to: 'the' }],
        replaceAsYouType: true,
        replaceHyperlinks: true,
      },
      'en',
    );
    store.update({ replaceAsYouType: false });
    expect(store.correct('teh')).toBe('teh');
    expect(onChange).toHaveBeenLastCalledWith(
      expect.objectContaining({ replaceAsYouType: false }),
      'en',
    );
    store.remove('teh');
    expect(store.state().entries).toEqual(symbols);

    locale = 'ja';
    options = { entries: () => [{ from: 'ｘ', to: 'y' }], replaceHyperlinks: false };
    expect(store.state()).toEqual({
      entries: [{ from: 'ｘ', to: 'y' }],
      replaceAsYouType: false,
      replaceHyperlinks: false,
    });
  });
});
