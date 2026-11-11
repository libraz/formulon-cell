import { describe, expect, it, vi } from 'vitest';
import {
  type AutoCorrectOptions,
  AutoCorrectStore,
  autoCorrectBefore,
  autoCorrectText,
  builtInAutoCorrectEntries,
  capitalizeDays,
  capitalizeSentences,
  correctTwoInitialCapitals,
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

  // The capitalisation rules are not measured against the desktop yet.
  it('lowers the second of two initial capitals, except plural acronyms and exceptions', () => {
    expect(correctTwoInitialCapitals('THe CApitals TWO Ok')).toBe('The Capitals TWO Ok');
    // A plural acronym keeps its capitals by rule, unmeasured; listed exceptions stay too.
    expect(correctTwoInitialCapitals('IDs PCs')).toBe('IDs PCs');
    expect(correctTwoInitialCapitals('MHz CDrom', ['cdrom'])).toBe('Mhz CDrom');
  });

  it('capitalises English day names in lower case, whatever the language', () => {
    expect(capitalizeDays('monday, Friday and sundays')).toBe('Monday, Friday and sundays');
  });

  it('capitalises a sentence after a sentence end and leaves the cell start as typed', () => {
    expect(capitalizeSentences('done. next! why? ok.no')).toBe('done. Next! Why? Ok.no');
    expect(capitalizeSentences('approx. ten. Approx. two', ['approx.'])).toBe(
      'approx. ten. Approx. two',
    );
    expect(capitalizeSentences('approx! ten', ['approx'])).toBe('approx! Ten');
  });

  it('applies each switch on its own', () => {
    let options: AutoCorrectOptions = {};
    const store = new AutoCorrectStore(
      () => options,
      () => 'ja',
    );
    expect(store.correct('THe end. monday')).toBe('The end. Monday');
    options = { correctTwoInitialCapitals: false, capitalizeSentences: false };
    expect(store.correct('THe end. monday')).toBe('THe end. Monday');
    store.update({ capitalizeDays: false });
    expect(store.correct('THe end. monday')).toBe('THe end. monday');
  });

  it('takes the host exceptions per language and hands edited ones back', () => {
    const onChange = vi.fn();
    const store = new AutoCorrectStore(
      () => ({
        exceptions: (locale) =>
          locale === 'en' ? { firstLetter: ['approx'], initialCaps: ['CDrom'] } : undefined,
        onChange,
      }),
      () => 'en',
    );
    expect(store.correct('approx. ten CDrom')).toBe('approx. ten CDrom');
    const exceptions = { firstLetter: [], initialCaps: [] };
    store.update({ exceptions });
    expect(store.correct('approx. ten CDrom')).toBe('approx. Ten Cdrom');
    expect(onChange).toHaveBeenLastCalledWith(expect.objectContaining({ exceptions }), 'en');
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
      exceptions: { firstLetter: [], initialCaps: [] },
      replaceAsYouType: true,
      replaceHyperlinks: true,
      correctTwoInitialCapitals: true,
      capitalizeSentences: true,
      capitalizeDays: true,
    });
    store.add('teh', 'the');
    expect(store.correct('teh (c)')).toBe('the ©');
    expect(onChange).toHaveBeenLastCalledWith(
      {
        entries: [...symbols, { from: 'teh', to: 'the' }],
        exceptions: { firstLetter: [], initialCaps: [] },
        replaceAsYouType: true,
        replaceHyperlinks: true,
        correctTwoInitialCapitals: true,
        capitalizeSentences: true,
        capitalizeDays: true,
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
      exceptions: { firstLetter: [], initialCaps: [] },
      replaceAsYouType: false,
      replaceHyperlinks: false,
      correctTwoInitialCapitals: true,
      capitalizeSentences: true,
      capitalizeDays: true,
    });
  });
});
