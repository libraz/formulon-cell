import type { SpreadsheetProfileId } from './types.js';
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
 * (`+-*\/^=<>!:&%@`), string quotes, parens, commas, semicolons, and array
 * column separators (`\\`).
 */
const IDENT_RE = /([^\s,;\\()+\-*/^=<>!:&%@"'][^\s,;\\()+\-*/^=<>!:&%@"']*)(?=\s*\()/g;

/** Locale ordinal mirror: 0 = en-US, 1 = ja-JP, 2 = de-DE, 3 = fr-FR. */
export type LocaleOrdinal = 0 | 1 | 2 | 3;

/**
 * Convert the legacy catalog ordinal into the engine's public profile while
 * retaining the active host (Windows or Mac). The old function-name adapter
 * accepted an ordinal; formulon 0.13 accepts an ExcelProfileId instead.
 */
export function spreadsheetProfileForLocaleOrdinal(
  activeProfile: SpreadsheetProfileId | null,
  locale: LocaleOrdinal,
): SpreadsheetProfileId {
  const host = activeProfile?.startsWith('mac-') ? 'mac' : 'windows';
  const suffix: Record<LocaleOrdinal, string> = {
    0: 'en_US',
    1: 'ja_JP',
    2: 'de_DE',
    3: 'fr_FR',
  };
  return `${host}-${suffix[locale]}` as SpreadsheetProfileId;
}

/** Resolve the function-name table used by a public spreadsheet profile. */
export function localeOrdinalForSpreadsheetProfile(
  profileId: SpreadsheetProfileId | null,
): LocaleOrdinal {
  if (profileId?.endsWith('-ja_JP')) return 1;
  if (profileId?.endsWith('-de_DE')) return 2;
  if (profileId?.endsWith('-fr_FR')) return 3;
  return 0;
}

/**
 * Walk every function-name token in `formula` (skipping string literals,
 * quoted sheet names, and structured-reference brackets) and remap it through
 * `mapName`. Used by both the
 * localize and canonicalize variants below — they only differ in which
 * engine method does the lookup.
 */
function remapFormulaIdents(formula: string, mapName: (raw: string) => string | null): string {
  if (!formula) return formula;
  // Build a list of replacement segments per syntax region so strings, quoted
  // sheet names, and structured references are preserved verbatim (commas /
  // parens inside any of those regions must not be treated as syntax).
  const segments: { start: number; end: number; insert: string }[] = [];
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
  const endOfQuotedRegion = (start: number, quote: '"' | "'"): number => {
    for (let i = start + 1; i < formula.length; i += 1) {
      if (formula[i] !== quote) continue;
      // Excel escapes a quote inside its corresponding quoted region by
      // doubling it (`""` in a string, `''` in a sheet name).
      if (formula[i + 1] === quote) {
        i += 1;
        continue;
      }
      return i + 1;
    }
    return formula.length;
  };
  const endOfBracketedRegion = (start: number): number => {
    let depth = 1;
    for (let i = start + 1; i < formula.length; i += 1) {
      if (formula[i] === '[') depth += 1;
      else if (formula[i] === ']') {
        depth -= 1;
        if (depth === 0) return i + 1;
      }
    }
    return formula.length;
  };
  let segmentStart = 0;
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch !== '"' && ch !== "'" && ch !== '[') {
      i += 1;
      continue;
    }
    considerSegment(segmentStart, i);
    i = ch === '[' ? endOfBracketedRegion(i) : endOfQuotedRegion(i, ch);
    segmentStart = i;
  }
  considerSegment(segmentStart, formula.length);
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
 * canonical form). The zh-CN, ko-KR, and th-TH profiles intentionally use
 * ordinal 0 because their function names are canonical English names. The
 * helper only changes display/editing copy; callers must keep canonical text
 * when persisting a formula.
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

/**
 * Canonicalize formula text at an engine write/evaluation boundary while
 * preserving names that are already canonical. Some locale tables contain an
 * alias which is also the canonical spelling of another function (Japanese
 * `DOLLAR`/`USDOLLAR` is the measured example). The public alias-only helper
 * intentionally keeps its historical lookup behavior; this boundary helper
 * consults the engine's canonical function registry before asking the locale
 * table to translate a token.
 */
export function canonicalizeFormulaForEngine(
  wb: WorkbookHandle,
  formula: string,
  locale: LocaleOrdinal,
): string {
  if (locale === 0) return formula;
  if (!wb.capabilities.functionLocale) return formula;

  let canonicalNames: ReadonlySet<string> | null = null;
  if (wb.capabilities.functionMetadata && typeof wb.functionNames === 'function') {
    try {
      const names = wb.functionNames();
      if (names) canonicalNames = new Set(names.map((name) => name.toUpperCase()));
    } catch {
      // An older or transient engine may not provide the catalog. Keep the
      // legacy alias lookup as the compatibility fallback in that case.
    }
  }

  return remapFormulaIdents(formula, (raw) => {
    if (canonicalNames?.has(raw.toUpperCase())) return null;
    const canonical = wb.canonicalizeFunctionName(raw, locale);
    if (!canonical || canonical === raw) return null;
    return canonical;
  });
}
