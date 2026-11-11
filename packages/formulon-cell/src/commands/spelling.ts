/** A spell checker the host supplies for Review > Spelling. `locale` is the spreadsheet's UI
 *  language, such as `'en'` or `'ja'`. */
export interface SpellingService {
  /** True when `word` is spelled correctly. */
  check(word: string, locale: string): boolean;
  /** Replacements for a misspelled word, best first. */
  suggest(word: string, locale: string): readonly string[];
  /** Called when the user adds a word, so the host can keep it beyond this session. */
  addWord?(word: string, locale: string): void;
}

const COMMON_TYPOS: Readonly<Record<string, string>> = {
  accomodate: 'accommodate',
  adress: 'address',
  definately: 'definitely',
  occured: 'occurred',
  recieve: 'receive',
  seperate: 'separate',
  teh: 'the',
  wierd: 'weird',
};

/** Gives `word` the capitalisation of `like`: all caps, a leading capital, or as is. */
const matchCase = (word: string, like: string): string => {
  if (like.length > 1 && like === like.toUpperCase()) return word.toUpperCase();
  const first = like.charAt(0);
  if (first !== first.toLowerCase()) return word.charAt(0).toUpperCase() + word.slice(1);
  return word;
};

const coveredLocale = (locale: string): boolean => locale.toLowerCase().startsWith('en');

const typoFix = (word: string): string | undefined => {
  const key = word.toLowerCase();
  return Object.hasOwn(COMMON_TYPOS, key) ? COMMON_TYPOS[key] : undefined;
};

/** The built-in checker: a short list of common English typos. Every word of another language
 *  passes, so it never flags what it cannot judge. */
export const defaultSpellingService: SpellingService = {
  check: (word, locale) => !coveredLocale(locale) || typoFix(word) === undefined,
  suggest: (word, locale) => {
    const fix = coveredLocale(locale) ? typoFix(word) : undefined;
    return fix === undefined ? [] : [matchCase(fix, word)];
  },
};

/** One word of a text, with its UTF-16 offsets. */
export interface SpellingWord {
  readonly word: string;
  readonly start: number;
  readonly end: number;
}

const WORD = /[\p{L}\p{M}\p{N}]+(?:['’][\p{L}\p{M}\p{N}]+)*/gu;
const DIGIT = /\p{N}/u;
const ADDRESS = /^(?:[a-z][a-z\d+.-]*:\/\/|www\.|\\\\[^\s\\])|^[^\s@]+@[^\s@]+\.[^\s@]+$/i;

/** True when the white-space-delimited `token` is an internet address, an e-mail address or a
 *  network path. */
export const isAddress = (token: string): boolean => ADDRESS.test(token);

/** The words of `text` a spelling check reads. Internet and e-mail addresses are skipped
 *  whole. */
export function spellingWords(text: string): SpellingWord[] {
  const words: SpellingWord[] = [];
  for (const chunk of text.matchAll(/\S+/g)) {
    if (isAddress(chunk[0])) continue;
    for (const match of chunk[0].matchAll(WORD)) {
      const start = (chunk.index ?? 0) + (match.index ?? 0);
      words.push({ word: match[0], start, end: start + match[0].length });
    }
  }
  return words;
}

/** Why a word stops the check. */
export type SpellingIssueKind = 'notInDictionary' | 'repeated';

/** What the check skips; the dialog's check boxes and the session's lists. */
export interface SpellingRules {
  readonly service: SpellingService;
  readonly locale: string;
  readonly ignoreUppercase: boolean;
  /** Words the user ignored for the session or added, lower-cased. */
  readonly accepted: ReadonlySet<string>;
}

const isUppercase = (word: string): boolean =>
  word !== word.toLowerCase() && word === word.toUpperCase();

/** The first word at or after `from` that stops the check, or null. A word that repeats the one
 *  before it, with only white space between, stops it as `repeated`. Words with digits pass. */
export function firstSpellingIssue(
  text: string,
  from: number,
  rules: SpellingRules,
): { readonly kind: SpellingIssueKind; readonly word: SpellingWord } | null {
  const words = spellingWords(text);
  for (let i = 0; i < words.length; i += 1) {
    const word = words[i] as SpellingWord;
    if (word.start < from) continue;
    if (DIGIT.test(word.word)) continue;
    const previous = words[i - 1];
    if (
      previous &&
      previous.word.toLowerCase() === word.word.toLowerCase() &&
      text.slice(previous.end, word.start).trim() === ''
    ) {
      return { kind: 'repeated', word };
    }
    if (rules.ignoreUppercase && isUppercase(word.word)) continue;
    if (rules.accepted.has(word.word.toLowerCase())) continue;
    if (!rules.service.check(word.word, rules.locale)) return { kind: 'notInDictionary', word };
  }
  return null;
}

/** `text` with every whole-word, same-case `word` replaced by `replacement`. */
export function replaceWholeWord(text: string, word: string, replacement: string): string {
  let out = '';
  let last = 0;
  for (const found of spellingWords(text)) {
    if (found.word !== word) continue;
    out += text.slice(last, found.start) + replacement;
    last = found.end;
  }
  return out + text.slice(last);
}

/** The text with the word at `[start, end)` replaced. Deleting a repeated word also drops the
 *  white space before it. */
export function replaceSpellingWord(
  text: string,
  word: SpellingWord,
  replacement: string,
  kind: SpellingIssueKind,
): string {
  let start = word.start;
  if (kind === 'repeated' && replacement === '') {
    while (start > 0 && /\s/.test(text.charAt(start - 1))) start -= 1;
  }
  return text.slice(0, start) + replacement + text.slice(word.end);
}
