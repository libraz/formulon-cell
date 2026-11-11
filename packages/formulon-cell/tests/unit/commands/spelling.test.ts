import { describe, expect, it } from 'vitest';
import {
  defaultSpellingService,
  firstSpellingIssue,
  replaceSpellingWord,
  replaceWholeWord,
  type SpellingRules,
  spellingWords,
} from '../../../src/commands/spelling.js';

const rules = (overrides: Partial<SpellingRules> = {}): SpellingRules => ({
  service: defaultSpellingService,
  locale: 'en',
  ignoreUppercase: true,
  accepted: new Set(),
  ...overrides,
});

describe('spelling', () => {
  it('flags the built-in English typos and passes every word of other languages', () => {
    expect(defaultSpellingService.check('teh', 'en')).toBe(false);
    expect(defaultSpellingService.check('the', 'en')).toBe(true);
    expect(defaultSpellingService.check('teh', 'ja')).toBe(true);
    expect(defaultSpellingService.suggest('Teh', 'en')).toEqual(['The']);
    expect(defaultSpellingService.suggest('RECIEVE', 'en')).toEqual(['RECEIVE']);
    expect(defaultSpellingService.suggest('teh', 'de')).toEqual([]);
  });

  it('reads words with their offsets and skips internet and e-mail addresses', () => {
    expect(spellingWords("don't see https://teh.example a@teh.example teh")).toEqual([
      { word: "don't", start: 0, end: 5 },
      { word: 'see', start: 6, end: 9 },
      { word: 'teh', start: 44, end: 47 },
    ]);
  });

  it('stops at a word not in the dictionary, from the given offset', () => {
    expect(firstSpellingIssue('teh cat teh', 0, rules())?.word.start).toBe(0);
    expect(firstSpellingIssue('teh cat teh', 1, rules())?.word.start).toBe(8);
    expect(firstSpellingIssue('teh cat', 3, rules())).toBeNull();
  });

  it('stops at a repeated word, and skips words with digits, accepted words and capitals', () => {
    expect(firstSpellingIssue('see the The end', 0, rules())).toEqual({
      kind: 'repeated',
      word: { word: 'The', start: 8, end: 11 },
    });
    expect(firstSpellingIssue('teh1 TEH', 0, rules())).toBeNull();
    expect(firstSpellingIssue('TEH', 0, rules({ ignoreUppercase: false }))?.kind).toBe(
      'notInDictionary',
    );
    expect(firstSpellingIssue('teh', 0, rules({ accepted: new Set(['teh']) }))).toBeNull();
  });

  it('replaces one word, or every same-case whole word', () => {
    const [word] = spellingWords('teh tehx teh');
    if (!word) throw new Error('no word');
    expect(replaceSpellingWord('teh tehx teh', word, 'the', 'notInDictionary')).toBe(
      'the tehx teh',
    );
    expect(replaceWholeWord('teh tehx Teh teh', 'teh', 'the')).toBe('the tehx Teh the');
  });

  it('deletes a repeated word with the space before it', () => {
    const issue = firstSpellingIssue('the  the end', 0, rules());
    if (issue?.kind !== 'repeated') throw new Error('no repeated word');
    expect(replaceSpellingWord('the  the end', issue.word, '', 'repeated')).toBe('the end');
  });
});
