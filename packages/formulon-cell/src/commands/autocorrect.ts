import type { SpreadsheetStore } from '../store/store.js';

/** One AutoCorrect replacement: typing `from` enters `to`. */
export interface AutoCorrectEntry {
  readonly from: string;
  readonly to: string;
}

/** What the AutoCorrect options hold for one language. */
export interface AutoCorrectState {
  readonly entries: readonly AutoCorrectEntry[];
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

type AutoCorrectSwitch = Exclude<keyof AutoCorrectState, 'entries'>;
const SWITCHES: readonly AutoCorrectSwitch[] = [
  'replaceAsYouType',
  'replaceHyperlinks',
  'correctTwoInitialCapitals',
  'capitalizeSentences',
  'capitalizeDays',
];

/** The host's part in AutoCorrect: a list per language, the starting switches, and a callback to
 *  keep what the user changes. `locale` is the spreadsheet's UI language. */
export interface AutoCorrectOptions {
  /** The list for `locale`, in place of the built-in one; `undefined` keeps the built-in one. */
  readonly entries?: (locale: string) => readonly AutoCorrectEntry[] | undefined;
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
  /** Called with the whole state after the user changes the list or a switch. */
  readonly onChange?: (state: AutoCorrectState, locale: string) => void;
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

/** The built-in list for `locale`. */
export const builtInAutoCorrectEntries = (locale: string): readonly AutoCorrectEntry[] =>
  SEED[locale.toLowerCase().split(/[-_]/)[0] ?? ''] ?? [];

const START = /\s/u;
const END = /[\s\p{P}\p{S}]/u;

/** Gives `word` the capitalisation of `like`: all caps, a leading capital, or as is. */
const carryCase = (word: string, like: string): string => {
  if (like.length > 1 && like === like.toUpperCase() && like !== like.toLowerCase()) {
    return word.toUpperCase();
  }
  const first = like.charAt(0);
  if (first !== first.toLowerCase()) return word.charAt(0).toUpperCase() + word.slice(1);
  return word;
};

/** The replacement when `entry` is typed at `start`, or null. A `from` without capitals matches
 *  any case and passes the typed capitalisation on; one with capitals matches exactly. */
const replacementAt = (text: string, start: number, entry: AutoCorrectEntry): string | null => {
  const typed = text.slice(start, start + entry.from.length);
  if (typed.length !== entry.from.length) return null;
  if (typed === entry.from) return entry.to;
  if (entry.from !== entry.from.toLowerCase() || typed.toLowerCase() !== entry.from) return null;
  return carryCase(entry.to, typed);
};

const startsWord = (text: string, at: number): boolean =>
  at === 0 || START.test(text.charAt(at - 1));
const endsWord = (text: string, at: number): boolean =>
  at === text.length || END.test(text.charAt(at));

const longestFirst = (entries: readonly AutoCorrectEntry[]): AutoCorrectEntry[] =>
  entries.filter((entry) => entry.from !== '').sort((a, b) => b.from.length - a.from.length);

/** The replacement of the entry typed just before `end`, the point where a delimiter was typed:
 *  what replacing as you type does. Null when none ends there. */
export function autoCorrectBefore(
  text: string,
  end: number,
  entries: readonly AutoCorrectEntry[],
): { readonly start: number; readonly end: number; readonly to: string } | null {
  if (!endsWord(text, end)) return null;
  for (const entry of longestFirst(entries)) {
    const start = end - entry.from.length;
    if (start < 0 || !startsWord(text, start)) continue;
    const to = replacementAt(text, start, entry);
    if (to !== null) return { start, end, to };
  }
  return null;
}

/** `text` with every entry replaced where it stands as a word: at the start or after white
 *  space, and before the end, white space or punctuation. */
export function autoCorrectText(text: string, entries: readonly AutoCorrectEntry[]): string {
  const ordered = longestFirst(entries);
  if (ordered.length === 0) return text;
  let out = '';
  let at = 0;
  while (at < text.length) {
    let replaced = false;
    if (startsWord(text, at)) {
      for (const entry of ordered) {
        const end = at + entry.from.length;
        if (!endsWord(text, end)) continue;
        const to = replacementAt(text, at, entry);
        if (to === null) continue;
        out += to;
        at = end;
        replaced = true;
        break;
      }
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

/** `text` with the second capital of a word like `THe` lowered: two capitals followed only by
 *  lower-case letters. */
export const correctTwoInitialCapitals = (text: string): string =>
  text.replace(WORD, (word) =>
    word.length > 2 && isUpper(word.slice(0, 2)) && isLower(word.slice(2))
      ? word.charAt(0) + word.charAt(1).toLowerCase() + word.slice(2)
      : word,
  );

/** `text` with English day names typed in lower case given a capital, in every language, as the
 *  desktop labels say. */
export const capitalizeDays = (text: string): string =>
  text.replace(WORD, (word) =>
    DAYS.has(word) ? word.charAt(0).toUpperCase() + word.slice(1) : word,
  );

/** `text` with the first letter after a sentence end and white space given a capital. The start
 *  of the cell is left as typed. */
export const capitalizeSentences = (text: string): string =>
  text.replace(
    /([.!?]\s+)(\p{Ll})/gu,
    (_, end: string, letter: string) => end + letter.toUpperCase(),
  );

/** The AutoCorrect state of one spreadsheet. User edits last for the instance's lifetime and are
 *  handed to the host through `onChange`. */
export class AutoCorrectStore {
  readonly #options: () => AutoCorrectOptions | undefined;
  readonly #locale: () => string;
  readonly #edited = new Map<string, readonly AutoCorrectEntry[]>();
  readonly #switches: Partial<Record<AutoCorrectSwitch, boolean>> = {};

  constructor(options: () => AutoCorrectOptions | undefined, locale: () => string) {
    this.#options = options;
    this.#locale = locale;
  }

  state(): AutoCorrectState {
    const locale = this.#locale();
    const options = this.#options();
    const on = (key: AutoCorrectSwitch): boolean => this.#switches[key] ?? options?.[key] ?? true;
    return {
      entries:
        this.#edited.get(locale) ?? options?.entries?.(locale) ?? builtInAutoCorrectEntries(locale),
      replaceAsYouType: on('replaceAsYouType'),
      replaceHyperlinks: on('replaceHyperlinks'),
      correctTwoInitialCapitals: on('correctTwoInitialCapitals'),
      capitalizeSentences: on('capitalizeSentences'),
      capitalizeDays: on('capitalizeDays'),
    };
  }

  /** Applies the user's change and tells the host. */
  update(patch: Partial<AutoCorrectState>): void {
    const locale = this.#locale();
    if (patch.entries) this.#edited.set(locale, [...patch.entries]);
    for (const key of SWITCHES) {
      const value = patch[key];
      if (value !== undefined) this.#switches[key] = value;
    }
    this.#options()?.onChange?.(this.state(), locale);
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

  /** Typed text as AutoCorrect leaves it on entry. */
  correct(text: string): string {
    const state = this.state();
    let out = state.replaceAsYouType ? autoCorrectText(text, state.entries) : text;
    if (state.correctTwoInitialCapitals) out = correctTwoInitialCapitals(out);
    if (state.capitalizeDays) out = capitalizeDays(out);
    if (state.capitalizeSentences) out = capitalizeSentences(out);
    return out;
  }
}

const stores = new WeakMap<SpreadsheetStore, AutoCorrectStore>();

export function registerAutoCorrect(store: SpreadsheetStore, autoCorrect: AutoCorrectStore): void {
  stores.set(store, autoCorrect);
}

export function autoCorrectFor(store: SpreadsheetStore): AutoCorrectStore | undefined {
  return stores.get(store);
}

/** A typed entry as AutoCorrect leaves it: constants are corrected, formulas left as typed. */
export const correctTypedEntry = (store: SpreadsheetStore, raw: string): string =>
  raw.trimStart().startsWith('=') ? raw : (stores.get(store)?.correct(raw) ?? raw);
