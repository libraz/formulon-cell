import { describe, expect, it } from 'vitest';
import {
  defaultSpellingService,
  firstSpellingIssue,
  type SpellingRules,
} from '../../../src/commands/spelling.js';

const rules: SpellingRules = {
  service: defaultSpellingService,
  locale: 'en',
  ignoreUppercase: true,
  accepted: new Set(),
};

describe('built-in spelling service', () => {
  it.each(['constructor', 'toString', '__proto__', 'hasOwnProperty'])(
    'passes the inherited object key %s without suggestions',
    (word) => {
      expect(defaultSpellingService.check(word, 'en')).toBe(true);
      expect(defaultSpellingService.suggest(word, 'en')).toEqual([]);
    },
  );
});

describe('firstSpellingIssue', () => {
  it('lets a repeated number pass as a word with digits', () => {
    expect(firstSpellingIssue('10 10', 0, rules)).toBeNull();
    expect(firstSpellingIssue('room1 room1', 0, rules)).toBeNull();
  });

  it('still flags a repeated word', () => {
    expect(firstSpellingIssue('the the', 0, rules)).toMatchObject({ kind: 'repeated' });
  });
});
