import { describe, expect, it } from 'vitest';
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

  /** A store over `initial` that keeps each committed configuration, as a host does. */
  const storeOver = (initial: AutoCorrectOptions, locale: () => string = () => 'en') => {
    const commits: AutoCorrectOptions[] = [];
    let options = initial;
    const store = new AutoCorrectStore(
      () => options,
      locale,
      (next) => {
        options = next;
        commits.push(next);
      },
    );
    return { store, commits, set: (next: AutoCorrectOptions) => (options = next) };
  };

  it('applies each switch on its own', () => {
    const { store, set } = storeOver({}, () => 'ja');
    expect(store.correct('THe end. monday')).toBe('The end. Monday');
    set({ correctTwoInitialCapitals: false, capitalizeSentences: false });
    expect(store.correct('THe end. monday')).toBe('THe end. Monday');
    store.update({ capitalizeDays: false });
    expect(store.correct('THe end. monday')).toBe('THe end. monday');
  });

  it('takes the exceptions per language and commits edited ones', () => {
    const { store, commits } = storeOver({
      exceptions: { en: { firstLetter: ['approx'], initialCaps: ['CDrom'] } },
    });
    expect(store.correct('approx. ten CDrom')).toBe('approx. ten CDrom');
    const exceptions = { firstLetter: [], initialCaps: [] };
    store.update({ exceptions });
    expect(store.correct('approx. ten CDrom')).toBe('approx. Ten Cdrom');
    expect(commits.at(-1)).toEqual({ exceptions: { en: exceptions } });
  });

  it('commits every edit as a new configuration, keyed by language', () => {
    let locale = 'en';
    const { store, commits, set } = storeOver({ replaceHyperlinks: false }, () => locale);
    expect(store.state()).toEqual({
      entries: symbols,
      exceptions: { firstLetter: [], initialCaps: [] },
      replaceAsYouType: true,
      replaceHyperlinks: false,
      correctTwoInitialCapitals: true,
      capitalizeSentences: true,
      capitalizeDays: true,
    });
    store.add('teh', 'the');
    expect(store.correct('teh (c)')).toBe('the ©');
    expect(commits.at(-1)).toEqual({
      replaceHyperlinks: false,
      entries: { en: [...symbols, { from: 'teh', to: 'the' }] },
    });
    store.update({ replaceAsYouType: false });
    expect(store.correct('teh')).toBe('teh');
    expect(commits.at(-1)).toMatchObject({ replaceAsYouType: false, replaceHyperlinks: false });
    store.remove('teh');
    expect(store.state().entries).toEqual(symbols);

    // Another language keeps its own list and the built-in one until edited.
    locale = 'ja';
    expect(store.state().entries).toEqual(builtInAutoCorrectEntries('ja'));
    set({ entries: { ja: [{ from: 'ｘ', to: 'y' }] } });
    expect(store.state().entries).toEqual([{ from: 'ｘ', to: 'y' }]);
  });
});
