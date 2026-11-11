import type { SpreadsheetStore } from '../store/store.js';
import { carryCase, isAddress } from './spelling.js';

/** One AutoCorrect replacement: typing `from` enters `to`. */
export interface AutoCorrectEntry {
  readonly from: string;
  readonly to: string;
}

/** Words the capitalisation rules leave alone. Matching ignores case. */
export interface AutoCorrectExceptions {
  /** Abbreviations after whose period a sentence does not start, such as `approx`; a trailing
   *  period is optional. */
  readonly firstLetter: readonly string[];
  /** Words kept with two initial capitals, such as `IDs`. */
  readonly initialCaps: readonly string[];
}

/** What the AutoCorrect options hold for one language. */
export interface AutoCorrectState {
  readonly entries: readonly AutoCorrectEntry[];
  readonly exceptions: AutoCorrectExceptions;
  /** Replace text as you type. */
  readonly replaceAsYouType: boolean;
  /** Replace internet and network paths with hyperlinks. */
  readonly replaceHyperlinks: boolean;
  /** Correct TWo INitial CApitals. */
  readonly correctTwoInitialCapitals: boolean;
  /** Capitalize first letter of sentences. */
  readonly capitalizeSentences: boolean;
  /** Capitalize names of days. */
  readonly capitalizeDays: boolean;
}

type AutoCorrectSwitch = Exclude<keyof AutoCorrectState, 'entries' | 'exceptions'>;
const SWITCHES: readonly AutoCorrectSwitch[] = [
  'replaceAsYouType',
  'replaceHyperlinks',
  'correctTwoInitialCapitals',
  'capitalizeSentences',
  'capitalizeDays',
];

/** AutoCorrect configuration: the lists per UI language and the switches. A key is a language
 *  (`'ja'`) or a full locale (`'ja-JP'`); the full locale wins over its language, and the user's
 *  edits are written under the locale `instance.i18n.locale` reports. The user's edits in AutoCorrect Options write a new value
 *  here and report `update` with `source: 'user'`, so the host keeps it from `instance.options`. */
export interface AutoCorrectOptions {
  /** Replacement lists by language; a language not listed keeps its built-in list. */
  readonly entries?: Readonly<Record<string, readonly AutoCorrectEntry[]>>;
  /** Capitalisation exceptions by language; a language not listed has none beyond the rules'. */
  readonly exceptions?: Readonly<Record<string, AutoCorrectExceptions>>;
  /** Replace text as you type; on by default. */
  readonly replaceAsYouType?: boolean;
  /** Replace internet and network paths with hyperlinks; on by default. */
  readonly replaceHyperlinks?: boolean;
  /** Correct TWo INitial CApitals; on by default. */
  readonly correctTwoInitialCapitals?: boolean;
  /** Capitalize first letter of sentences; on by default. */
  readonly capitalizeSentences?: boolean;
  /** Capitalize names of days; on by default. */
  readonly capitalizeDays?: boolean;
}

const SYMBOLS: readonly AutoCorrectEntry[] = [
  { from: '(c)', to: '©' },
  { from: '(e)', to: '€' },
  { from: '(r)', to: '®' },
  { from: '(tm)', to: '™' },
  { from: '...', to: '…' },
];

/** The generic part of the desktop lists: their symbol replacements, and the two Japanese
 *  greetings. Korean and Thai start empty, as on the desktop. */
const SEED: Readonly<Record<string, readonly AutoCorrectEntry[]>> = {
  en: SYMBOLS,
  de: SYMBOLS,
  fr: SYMBOLS,
  zh: SYMBOLS.filter((entry) => entry.from !== '(e)'),
  ja: [
    { from: 'こんにちわ', to: 'こんにちは' },
    { from: 'こんばんわ', to: 'こんばんは' },
  ],
};

/** The value keyed by the full `locale`, else by its language; the full locale wins. */
const byLocale = <T>(
  record: Readonly<Record<string, T>> | undefined,
  locale: string,
): T | undefined => {
  if (!record) return undefined;
  if (Object.hasOwn(record, locale)) return record[locale];
  const language = locale.split(/[-_]/)[0] ?? '';
  return Object.hasOwn(record, language) ? record[language] : undefined;
};

/** The built-in list for `locale`. */
export const builtInAutoCorrectEntries = (locale: string): readonly AutoCorrectEntry[] =>
  SEED[locale.toLowerCase().split(/[-_]/)[0] ?? ''] ?? [];

const START = /\s/u;
const END = /[\s\p{P}\p{S}]/u;

/** The replacement when `entry` is typed at `start`, or null. A `from` without capitals matches
 *  any case and passes the typed capitalisation on; one with capitals matches exactly. */
const replacementAt = (text: string, start: number, entry: AutoCorrectEntry): string | null => {
  const typed = text.slice(start, start + entry.from.length);
  if (typed.length !== entry.from.length) return null;
  if (typed === entry.from) return entry.to;
  if (entry.from !== entry.from.toLowerCase() || typed.toLowerCase() !== entry.from) return null;
  return carryCase(entry.to, typed);
};

/** Han and kana, written without spaces between words. */
const CJK = /^[\p{scx=Han}\p{scx=Hiragana}\p{scx=Katakana}]+$/u;

const startsWord = (text: string, at: number): boolean =>
  at === 0 || START.test(text.charAt(at - 1));
const endsWord = (text: string, at: number): boolean =>
  at === text.length || END.test(text.charAt(at));

const longestFirst = (entries: readonly AutoCorrectEntry[]): AutoCorrectEntry[] =>
  entries.filter((entry) => entry.from !== '').sort((a, b) => b.from.length - a.from.length);

/** `text` with every entry replaced where it stands as a word: at the start or after white
 *  space, and before the end, white space or punctuation. A Han or kana entry needs no word
 *  boundary, as those scripts have no spaces between words. */
export function autoCorrectText(text: string, entries: readonly AutoCorrectEntry[]): string {
  const ordered = longestFirst(entries);
  if (ordered.length === 0) return text;
  let out = '';
  let at = 0;
  while (at < text.length) {
    let replaced = false;
    const wordStart = startsWord(text, at);
    for (const entry of ordered) {
      const end = at + entry.from.length;
      const free = CJK.test(entry.from);
      if (!free && (!wordStart || !endsWord(text, end))) continue;
      const to = replacementAt(text, at, entry);
      if (to === null) continue;
      out += to;
      at = end;
      replaced = true;
      break;
    }
    if (!replaced) {
      out += text.charAt(at);
      at += 1;
    }
  }
  return out;
}

const WORD = /\p{L}+/gu;
const DAYS = new Set([
  'monday',
  'tuesday',
  'wednesday',
  'thursday',
  'friday',
  'saturday',
  'sunday',
]);
const isLower = (text: string): boolean =>
  text === text.toLowerCase() && text !== text.toUpperCase();
const isUpper = (text: string): boolean =>
  text === text.toUpperCase() && text !== text.toLowerCase();

const NO_EXCEPTIONS: AutoCorrectExceptions = { firstLetter: [], initialCaps: [] };
const lowerSet = (words: readonly string[], strip = ''): Set<string> =>
  new Set(
    words.map((word) => (strip && word.endsWith(strip) ? word.slice(0, -1) : word).toLowerCase()),
  );

/** True when the white-space-delimited token around `index` is an address or path. */
const inAddress = (text: string, index: number): boolean => {
  let start = index;
  while (start > 0 && !/\s/.test(text.charAt(start - 1))) start -= 1;
  let end = index;
  while (end < text.length && !/\s/.test(text.charAt(end))) end += 1;
  return isAddress(text.slice(start, end));
};

/** `text` with the second capital of a word like `THe` lowered: two capitals followed only by
 *  lower-case letters. A plural acronym such as `PCs` stays as typed when a delimiter follows it,
 *  but not as the last word of a committed entry (`endsEntry`), as on the desktop; the
 *  `exceptions` always stay. */
export const correctTwoInitialCapitals = (
  text: string,
  exceptions: readonly string[] = [],
  endsEntry = true,
): string => {
  const keep = lowerSet(exceptions);
  return text.replace(WORD, (word, at: number) =>
    word.length > 2 &&
    isUpper(word.slice(0, 2)) &&
    isLower(word.slice(2)) &&
    (word.slice(2) !== 's' || (endsEntry && at + word.length === text.length)) &&
    !keep.has(word.toLowerCase()) &&
    !inAddress(text, at)
      ? word.charAt(0) + word.charAt(1).toLowerCase() + word.slice(2)
      : word,
  );
};

/** `text` with English day names typed in lower case given a capital, in every language, as the
 *  desktop labels say. */
export const capitalizeDays = (text: string): string =>
  text.replace(WORD, (word, at: number) =>
    DAYS.has(word) && !inAddress(text, at) ? word.charAt(0).toUpperCase() + word.slice(1) : word,
  );

/** `text` with the first letter after a sentence end and white space given a capital. The start
 *  of the cell, and the word after the period of one of the `abbreviations`, are left as typed. */
export const capitalizeSentences = (
  text: string,
  abbreviations: readonly string[] = [],
): string => {
  const keep = lowerSet(abbreviations, '.');
  return text.replace(
    /(\p{L}*)([.!?]\s+)(\p{Ll})/gu,
    (match, word: string, end: string, letter: string, at: number) =>
      (end.startsWith('.') && keep.has(word.toLowerCase())) ||
      inAddress(text, at + match.length - 1)
        ? match
        : word + end + letter.toUpperCase(),
  );
};

/** The AutoCorrect state of one spreadsheet, read from its configuration. A user edit writes a new
 *  configuration through `commit`. */
export class AutoCorrectStore {
  readonly #options: () => AutoCorrectOptions | undefined;
  readonly #locale: () => string;
  readonly #commit: (next: AutoCorrectOptions) => void;

  constructor(
    options: () => AutoCorrectOptions | undefined,
    locale: () => string,
    commit: (next: AutoCorrectOptions) => void,
  ) {
    this.#options = options;
    this.#locale = locale;
    this.#commit = commit;
  }

  state(): AutoCorrectState {
    const locale = this.#locale();
    const options = this.#options();
    const on = (key: AutoCorrectSwitch): boolean => options?.[key] ?? true;
    return {
      entries: byLocale(options?.entries, locale) ?? builtInAutoCorrectEntries(locale),
      exceptions: byLocale(options?.exceptions, locale) ?? NO_EXCEPTIONS,
      replaceAsYouType: on('replaceAsYouType'),
      replaceHyperlinks: on('replaceHyperlinks'),
      correctTwoInitialCapitals: on('correctTwoInitialCapitals'),
      capitalizeSentences: on('capitalizeSentences'),
      capitalizeDays: on('capitalizeDays'),
    };
  }

  /** Applies the user's change as a new configuration. */
  update(patch: Partial<AutoCorrectState>): void {
    const locale = this.#locale();
    const options = this.#options() ?? {};
    const switches = Object.fromEntries(
      SWITCHES.flatMap((key) => (patch[key] === undefined ? [] : [[key, patch[key]]])),
    ) as Partial<Record<AutoCorrectSwitch, boolean>>;
    this.#commit({
      ...options,
      ...switches,
      ...(patch.entries ? { entries: { ...options.entries, [locale]: [...patch.entries] } } : {}),
      ...(patch.exceptions
        ? { exceptions: { ...options.exceptions, [locale]: patch.exceptions } }
        : {}),
    });
  }

  /** Adds `from` → `to`, replacing an entry with the same `from`. */
  add(from: string, to: string): void {
    const entries = this.state().entries.filter((entry) => entry.from !== from);
    this.update({ entries: [...entries, { from, to }] });
  }

  /** Removes the entry for `from`. */
  remove(from: string): void {
    this.update({ entries: this.state().entries.filter((entry) => entry.from !== from) });
  }

  /** Typed text as AutoCorrect leaves it on entry. `endsEntry` is false for text a delimiter
   *  follows, such as the words before the caret while typing. */
  correct(text: string, endsEntry = true): string {
    const state = this.state();
    let out = state.replaceAsYouType ? autoCorrectText(text, state.entries) : text;
    if (state.correctTwoInitialCapitals) {
      out = correctTwoInitialCapitals(out, state.exceptions.initialCaps, endsEntry);
    }
    if (state.capitalizeDays) out = capitalizeDays(out);
    if (state.capitalizeSentences) {
      out = capitalizeSentences(out, state.exceptions.firstLetter);
    }
    return out;
  }

  /** A typed entry as AutoCorrect leaves it: a constant whole, a formula inside its strings. */
  correctEntry(raw: string, endsEntry = true): string {
    if (!raw.trimStart().startsWith('=')) return this.correct(raw, endsEntry);
    // A closed string has its quote after the last word; an open one ends where the entry does.
    return raw.replace(
      /"((?:[^"]|"")*)("|$)/g,
      (_match, inner: string, close: string) =>
        `"${this.correct(inner, endsEntry && close === '')}${close}`,
    );
  }

  /** Replacing as you type: when the character before `caret` ends a word, the entry with the
   *  text up to it corrected and the caret after the correction. Null when nothing changes. */
  live(value: string, caret: number): { readonly value: string; readonly caret: number } | null {
    if (caret === 0 || !END.test(value.charAt(caret - 1))) return null;
    const head = value.slice(0, caret);
    const fixed = this.correctEntry(head, false);
    return fixed === head ? null : { value: fixed + value.slice(caret), caret: fixed.length };
  }
}

const stores = new WeakMap<SpreadsheetStore, AutoCorrectStore>();

export function registerAutoCorrect(store: SpreadsheetStore, autoCorrect: AutoCorrectStore): void {
  stores.set(store, autoCorrect);
}

export function autoCorrectFor(store: SpreadsheetStore): AutoCorrectStore | undefined {
  return stores.get(store);
}

/** A typed entry as AutoCorrect leaves it on commit: the last word gets its correction here,
 *  the words before it already had theirs at their delimiters. */
export const correctTypedEntry = (store: SpreadsheetStore, raw: string): string =>
  stores.get(store)?.correctEntry(raw) ?? raw;

/** Replacing as you type in an entry field: corrects `input` in place after a typed delimiter. */
export function autoCorrectInput(
  store: SpreadsheetStore,
  input: HTMLInputElement | HTMLTextAreaElement,
): void {
  const caret = input.selectionStart;
  if (caret === null || caret !== input.selectionEnd) return;
  const fixed = stores.get(store)?.live(input.value, caret);
  if (!fixed) return;
  input.value = fixed.value;
  input.setSelectionRange(fixed.caret, fixed.caret);
}
