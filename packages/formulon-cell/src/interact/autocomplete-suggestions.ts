import { formulaCaretContext, formulaIdentifierPrefix } from '../commands/formula-edit-context.js';
import type { AutocompleteName, AutocompleteTable } from './autocomplete-names.js';

export type AutocompleteSuggestionKind = 'function' | 'tableColumn' | 'column' | 'history' | 'name';

export interface SuggestionContext {
  matches: string[];
  /** Character offset where the partial token starts (inclusive). */
  tokenStart: number;
  /** Character offset where the partial token ends (exclusive). */
  tokenEnd: number;
  /** What gets inserted when the user accepts a match — defaults to `${match}`
   *  for refs and `${match}(` for functions. */
  insertSuffix: string;
  /** The kind drives display labels and badges. */
  kind: AutocompleteSuggestionKind;
  /** Per-match kind and detail when the list mixes functions and names. */
  kinds?: AutocompleteSuggestionKind[];
  details?: string[];
  /** Candidate-specific replacement text, used when raw labels need syntax
   *  encoding before they are inserted. */
  insertions?: string[];
  /** Number of existing characters after the replacement that place the caret
   *  inside a call parenthesis already present in the formula. */
  existingCallParenLength?: number;
  /** Headed groups, in list order, when the list is grouped. */
  groups?: AutocompleteGroup[];
}

/** A headed run of the list from `start` up to the next group. */
export interface AutocompleteGroup {
  start: number;
  kind: 'name' | 'table' | 'function';
}

/** Suggestions replace the token prefix but leave the suffix in place, so a caret that splits a
 *  token offers nothing: accepting would duplicate the token's remaining text. */
export function caretSplitsToken(text: string, caret: number): boolean {
  const next = text[caret] ?? '';
  return next !== '' && /[\p{L}\p{N}_.\\]/u.test(next);
}

export function existingFunctionParenLength(text: string, caret: number): number | null {
  const context = formulaCaretContext(text, caret);
  if (context !== 'function' && context !== 'expression') return null;
  const { token } = formulaIdentifierPrefix(text, caret);
  if (!/^[\p{L}_]/u.test(token)) return null;
  const openParen = /^\s*\(/u.exec(text.slice(caret))?.[0];
  if (!openParen) return null;
  if (context === 'function' || openParen.length > 1) return openParen.length;
  return null;
}

/** column-history autocomplete: when editing a plain-text cell, match
 *  what the user has typed so far against the values already entered above in
 *  the same column. `values` must already be deduped, blanks excluded, in
 *  nearest-first order (caller's responsibility). The whole input acts as the
 *  partial token — desktop spreadsheets replaces the entire cell text on accept, never just
 *  a fragment. Returns null when the input is empty or no value prefix-matches. */
export function suggestColumnHistory(
  text: string,
  caret: number,
  values: readonly string[],
): SuggestionContext | null {
  // Only suggest when caret is at end and the input is a non-empty token. A
  // shorter caret (mid-edit) means the user is correcting earlier characters,
  // not extending the tail — spreadsheets don't pop the list there either.
  if (caret !== text.length) return null;
  if (text.length === 0) return null;
  const lower = text.toLowerCase();
  const matches: string[] = [];
  for (const v of values) {
    if (v.length === text.length) continue; // skip exact-length (would be no-op)
    if (v.toLowerCase().startsWith(lower)) matches.push(v);
    if (matches.length >= 10) break;
  }
  if (matches.length === 0) return null;
  return {
    matches,
    tokenStart: 0,
    tokenEnd: text.length,
    insertSuffix: '',
    kind: 'column',
  };
}

/** Filter the consumer's custom function list against the partial token at
 *  `caret`. Mirrors `suggestFunctions` but against a runtime-supplied list. */
export function suggestCustomFunctions(
  text: string,
  caret: number,
  names: readonly string[],
): { token: string; tokenStart: number; matches: string[] } | null {
  if (!text.startsWith('=') || names.length === 0) return null;
  const caretContext = formulaCaretContext(text, caret);
  if (caretContext !== 'expression' && caretContext !== 'function') return null;
  const { token, tokenStart } = formulaIdentifierPrefix(text, caret);
  if (token.length < 1 || !/^[\p{L}_]/u.test(token)) return null;
  const upper = token.toLocaleUpperCase();
  const matches = names.filter((n) => startsWithFolded(n, upper));
  if (matches.length === 0) return null;
  return { token, tokenStart, matches };
}

/** Case-insensitive prefix test; `upperToken` is already upper-cased. */
const startsWithFolded = (label: string, upperToken: string): boolean =>
  label.toLocaleUpperCase().startsWith(upperToken);

/** The items whose label starts with `upperToken` first, the rest after, each keeping its order. */
export const partitionByPrefix = <T>(
  items: readonly T[],
  label: (item: T) => string,
  upperToken: string,
): T[] => [
  ...items.filter((item) => startsWithFolded(label(item), upperToken)),
  ...items.filter((item) => !startsWithFolded(label(item), upperToken)),
];

/** Stable sort by lowercased label, lowercasing each label once. */
export const sortCaseless = <T>(items: readonly T[], label: (item: T) => string): T[] =>
  items
    .map((item) => ({ item, key: label(item).toLocaleLowerCase() }))
    .sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0))
    .map((d) => d.item);

/** Filter defined/table names against the partial token at `caret`. The match
 *  is a case-insensitive prefix, so Japanese names work the same way. */
export function suggestNames(
  text: string,
  caret: number,
  names: readonly AutocompleteName[],
): { tokenStart: number; matches: AutocompleteName[] } | null {
  if (!text.startsWith('=') || names.length === 0) return null;
  if (formulaCaretContext(text, caret) !== 'expression') return null;
  const { token, tokenStart } = formulaIdentifierPrefix(text, caret);
  if (token.length < 1 || !/^[\p{L}_\\]/u.test(token)) return null;
  const upper = token.toLocaleUpperCase();
  const seen = new Set<string>();
  const matches = names.filter((n) => {
    const key = n.name.toLocaleUpperCase();
    if (!key.startsWith(upper) || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  return matches.length === 0 ? null : { tokenStart, matches };
}

const STRUCTURED_REF_EXTRA_BRACKET_CHARS = new Set([
  '\t',
  '\n',
  '\r',
  ',',
  ':',
  '.',
  '[',
  ']',
  '#',
  "'",
  '"',
  '{',
  '}',
  '$',
  '^',
  '&',
  '*',
  '+',
  '=',
  '-',
  '>',
  '<',
  '/',
  '@',
  '\\',
  '!',
  '(',
  ')',
  '%',
  '?',
  '`',
  ';',
  '~',
  '_',
]);

const STRUCTURED_REF_ESCAPED_CHARS = new Set(['[', ']', '#', "'", '@']);

/** Offsets of the `[` still open at `caret`. Outside brackets, "..." and '...' spans are
 *  skipped; inside brackets, `'` escapes the next special character. */
function openStructuredRefBrackets(text: string, caret: number): number[] {
  const opens: number[] = [];
  let inString = false;
  let inSheet = false;
  let i = text.startsWith('=') ? 1 : 0;
  const stop = Math.max(i, Math.min(caret, text.length));

  while (i < stop) {
    const ch = text[i] ?? '';
    const next = text[i + 1] ?? '';
    if (inString) {
      if (ch === '"') {
        if (next === '"') {
          i += 2;
          continue;
        }
        inString = false;
      }
      i += 1;
      continue;
    }
    if (inSheet) {
      if (ch === "'") {
        if (next === "'") {
          i += 2;
          continue;
        }
        inSheet = false;
      }
      i += 1;
      continue;
    }
    if (opens.length > 0) {
      if (ch === "'" && STRUCTURED_REF_ESCAPED_CHARS.has(next)) {
        i += 2;
        continue;
      }
      if (ch === '[') opens.push(i);
      else if (ch === ']') opens.pop();
      i += 1;
      continue;
    }
    if (ch === '"') {
      inString = true;
      i += 1;
      continue;
    }
    if (ch === "'") {
      inSheet = true;
      i += 1;
      continue;
    }
    if (ch === '[') opens.push(i);
    i += 1;
  }
  return opens;
}

function decodeStructuredRefPartial(partial: string): string {
  let decoded = '';
  for (let i = 0; i < partial.length; i += 1) {
    const ch = partial[i] ?? '';
    const next = partial[i + 1] ?? '';
    // A dangling escape still identifies a useful prefix (`A'` before the escaped `]` in `A]B`).
    if (ch === "'" && i === partial.length - 1) continue;
    if (ch === "'" && STRUCTURED_REF_ESCAPED_CHARS.has(next)) {
      decoded += next;
      i += 1;
    } else {
      decoded += ch;
    }
  }
  return decoded;
}

function encodeStructuredRefColumn(header: string, openBracketCount: number): string {
  const escaped = [...header]
    .map((ch) => (STRUCTURED_REF_ESCAPED_CHARS.has(ch) ? `'${ch}` : ch))
    .join('');
  const needsExtraBracket = [...header].some((ch) => STRUCTURED_REF_EXTRA_BRACKET_CHARS.has(ch));
  const addInnerBracket = needsExtraBracket && openBracketCount === 1;
  const closeCount = openBracketCount + (addInnerBracket ? 1 : 0);
  return `${addInnerBracket ? '[' : ''}${escaped}${']'.repeat(closeCount)}`;
}

/** Detect a `TableName[partial` token immediately before the caret and return
 *  matching column names. The token is allowed to contain spaces inside the
 *  bracket so multi-word column names work. Returns null when the caret is
 *  outside such a token. */
export function suggestStructuredRef(
  text: string,
  caret: number,
  tables: readonly AutocompleteTable[],
): SuggestionContext | null {
  if (!text.startsWith('=') || tables.length === 0) return null;
  const caretContext = formulaCaretContext(text, caret);
  if (caretContext === 'string' || caretContext === 'sheet') return null;
  const next = text[caret] ?? '';
  if (next && /[\p{L}\p{N}_.\\\]]/u.test(next)) return null;
  const openBrackets = openStructuredRefBrackets(text, caret);
  if (openBrackets.length === 0) return null;
  const tableBracket = openBrackets[0];
  const columnBracket = openBrackets[openBrackets.length - 1];
  if (tableBracket === undefined || columnBracket === undefined) return null;
  const { token: tableName } = formulaIdentifierPrefix(text, tableBracket);
  if (tableName.length === 0 || !/^[\p{L}_\\]/u.test(tableName)) return null;
  const foldedTableName = tableName.toLocaleLowerCase();
  const table = tables.find((t) => t.name.toLocaleLowerCase() === foldedTableName);
  if (!table || table.columns.length === 0) return null;
  const partial = decodeStructuredRefPartial(text.slice(columnBracket + 1, caret));
  const lower = partial.toLocaleLowerCase();
  const seen = new Set<string>();
  const matches = sortCaseless(
    table.columns.filter((c) => {
      const key = c.toLocaleLowerCase();
      if (!key.startsWith(lower) || seen.has(key)) return false;
      seen.add(key);
      return true;
    }),
    (c) => c,
  );
  if (matches.length === 0) return null;
  return {
    matches,
    tokenStart: columnBracket + 1,
    tokenEnd: caret,
    insertSuffix: ']',
    kind: 'tableColumn',
    insertions: matches.map((header) => encodeStructuredRefColumn(header, openBrackets.length)),
  };
}
