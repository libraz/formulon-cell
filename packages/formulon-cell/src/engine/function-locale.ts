import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Per-call regex that matches every identifier that *could* be a function
 * name — anything followed by `(` that isn't a reserved delimiter. Built
 * to handle Unicode (hiragana / katakana / kanji function aliases) since
 * `\b` in JS only recognises ASCII word boundaries.
 *
 * Refs (A1, $B$2) are still walked but they're never followed by `(`, so
 * they don't enter the rename path. The caller's `mapName` returns the
 * input unchanged for unknown tokens, making accidental matches harmless.
 *
 * Excluded delimiters: whitespace, the operator set
 * (`+-*\/^=<>!:&%`), string quotes, parens, and commas.
 */
const IDENT_RE = /([^\s,()+\-*/^=<>!:&%"'][^\s,()+\-*/^=<>!:&%"']*)(?=\s*\()/g;

/** Locale ordinal mirror: 0 = en-US, 1 = ja-JP. */
export type LocaleOrdinal = 0 | 1;

/**
 * Walk every function-name token in `formula` (skipping the contents of
 * string literals) and remap it through `mapName`. Used by both the
 * localize and canonicalize variants below — they only differ in which
 * engine method does the lookup.
 */
function remapFormulaIdents(formula: string, mapName: (raw: string) => string | null): string {
  if (!formula) return formula;
  // Build a list of replacement segments per non-string region so string
  // literal contents are preserved verbatim (commas / parens inside a
  // string must not be treated as syntax).
  const segments: { start: number; end: number; insert: string }[] = [];
  let inString = false;
  let segmentStart = 0;
  const considerSegment = (segStart: number, segEnd: number): void => {
    const slice = formula.slice(segStart, segEnd);
    IDENT_RE.lastIndex = 0;
    let match = IDENT_RE.exec(slice);
    while (match !== null) {
      const raw = match[1];
      if (raw) {
        const mapped = mapName(raw);
        if (mapped && mapped !== raw) {
          segments.push({
            start: segStart + match.index,
            end: segStart + match.index + raw.length,
            insert: mapped,
          });
        }
      }
      match = IDENT_RE.exec(slice);
    }
  };
  for (let i = 0; i < formula.length; i += 1) {
    if (formula[i] === '"') {
      if (!inString) {
        considerSegment(segmentStart, i);
        inString = true;
      } else {
        inString = false;
        segmentStart = i + 1;
      }
    }
  }
  if (!inString && segmentStart < formula.length) {
    considerSegment(segmentStart, formula.length);
  }
  if (segments.length === 0) return formula;
  let out = '';
  let cursor = 0;
  for (const seg of segments) {
    out += formula.slice(cursor, seg.start);
    out += seg.insert;
    cursor = seg.end;
  }
  out += formula.slice(cursor);
  return out;
}

/**
 * Replace every canonical function name in `formula` with its localized
 * alias for `locale`. Returns the input unchanged when the engine doesn't
 * expose `localizeFunctionName` or when `locale === 0` (en-US is the
 * canonical form). Today, ja-JP returns the canonical name unchanged
 * upstream — this helper is wired up so cell will pick up aliases the
 * day formulon publishes them, without further consumer-side changes.
 */
export function localizeFormula(
  wb: WorkbookHandle,
  formula: string,
  locale: LocaleOrdinal,
): string {
  if (locale === 0) return formula;
  if (!wb.capabilities.functionLocale) return formula;
  return remapFormulaIdents(formula, (raw) => {
    const mapped = wb.localizeFunctionName(raw, locale);
    if (!mapped || mapped === raw) return null;
    return mapped;
  });
}

/**
 * Inverse of `localizeFormula`: translate any localized function names
 * back to the canonical en-US form before sending the text through the
 * parser. Idempotent for already-canonical formulas. Same fallback rules
 * as `localizeFormula`.
 */
export function canonicalizeFormula(
  wb: WorkbookHandle,
  formula: string,
  locale: LocaleOrdinal,
): string {
  if (locale === 0) return formula;
  if (!wb.capabilities.functionLocale) return formula;
  return remapFormulaIdents(formula, (raw) => {
    const canonical = wb.canonicalizeFunctionName(raw, locale);
    if (!canonical || canonical === raw) return null;
    return canonical;
  });
}
