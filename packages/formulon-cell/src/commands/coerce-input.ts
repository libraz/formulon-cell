import { effectiveCellFormatWithPending } from '../engine/cell-format-sync.js';
import { formatCodeToNumFmt } from '../engine/format-writeback.js';
import {
  canonicalizeFormula,
  localeOrdinalForSpreadsheetProfile,
} from '../engine/function-locale.js';
import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Addr, CellValue, SpreadsheetProfileId } from '../engine/types.js';
import { formatError, fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { spreadsheetProfileForLocale } from '../mount/locale-profile.js';
import { formatWithPending } from '../store/pending-format.js';
import {
  type CellValidation,
  mutators,
  type NumFmt,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { inputLanguage, parseNumberInput, parseTypedValue } from './input-grammar.js';
import { isCellWritable, warnProtected } from './protection.js';
import { extractRefs, normalizeR1C1Formula } from './refs.js';
import {
  type CustomValidationEvaluator,
  nativeValidationOutcome,
  type ValidationOutcome,
  validateAgainst,
} from './validate.js';

export type CoercedInput =
  | { kind: 'blank' }
  | { kind: 'formula'; text: string }
  | { kind: 'number'; value: number; implicitFormat?: NumFmt }
  | { kind: 'bool'; value: boolean }
  | { kind: 'error'; code: number; text: string }
  | { kind: 'text'; value: string };

export interface CoerceInputOptions {
  forceText?: boolean;
  /** BCP47 input locale. When omitted, the input grammar uses en-US. */
  locale?: string;
}

/** Boolean spellings typed input accepts, per input language. */
const BOOLEAN_TEXT: Readonly<Record<string, readonly [string, string]>> = {
  de: ['WAHR', 'FALSCH'],
  fr: ['VRAI', 'FAUX'],
};

const booleanText = (locale: string | undefined): readonly [string, string] =>
  BOOLEAN_TEXT[inputLanguage(locale)] ?? ['TRUE', 'FALSE'];

/** Error literals typed input stores as error values, by ErrorCode ordinal.
 *  German and French accept only their localized spellings. #GETTING_DATA is
 *  accepted but saved as #N/A. */
const ERROR_LITERALS: Readonly<Record<string, Readonly<Record<string, number>>>> = {
  en: {
    '#NULL!': 0,
    '#DIV/0!': 1,
    '#VALUE!': 2,
    '#REF!': 3,
    '#NAME?': 4,
    '#NUM!': 5,
    '#N/A': 6,
    '#GETTING_DATA': 6,
  },
  de: {
    '#NULL!': 0,
    '#DIV/0!': 1,
    '#WERT!': 2,
    '#BEZUG!': 3,
    '#NAME?': 4,
    '#ZAHL!': 5,
    '#NV': 6,
  },
  fr: {
    '#NUL!': 0,
    '#DIV/0!': 1,
    '#VALEUR!': 2,
    '#REF!': 3,
    '#NOM?': 4,
    '#NOMBRE!': 5,
    '#N/A': 6,
  },
};

const errorLiteral = (text: string, locale: string | undefined): number | null => {
  const literals = ERROR_LITERALS[inputLanguage(locale)] ?? ERROR_LITERALS.en;
  return literals?.[text.toUpperCase()] ?? null;
};

const PROFILE_LOCALE_TO_INPUT_LOCALE: Readonly<Record<string, string>> = {
  '-ja_JP': 'ja-JP',
  '-en_US': 'en-US',
  '-de_DE': 'de-DE',
  '-fr_FR': 'fr-FR',
  '-zh_CN': 'zh-CN',
  '-zh_TW': 'zh-TW',
  '-ko_KR': 'ko-KR',
  '-th_TH': 'th-TH',
};

/** Resolve an explicit input locale to the public profile used by the live
 * workbook, keeping its current Windows/Mac platform. CSV import and paste
 * can request a locale without mutating the workbook profile. */
function publicProfileForInputLocale(
  wb: WorkbookHandle,
  locale: string | undefined,
): SpreadsheetProfileId | null | undefined {
  const current = typeof wb.spreadsheetProfileId === 'function' ? wb.spreadsheetProfileId() : null;
  if (locale === undefined) return current ?? undefined;
  return spreadsheetProfileForLocale(locale, current?.startsWith('mac-') ? 'mac' : 'default');
}

/** Convert user-facing localized function names to formulon's canonical
 * storage grammar. The optional adapter method is used when available; the
 * fallback keeps compatibility with the first profile API. */
export function canonicalizeLocalizedFormulaForInput(
  wb: WorkbookHandle,
  formula: string,
  locale?: string,
): string {
  const profile = publicProfileForInputLocale(wb, locale);
  if (profile === null) return formula;
  if (typeof wb.canonicalizeLocalizedFormula === 'function') {
    return wb.canonicalizeLocalizedFormula(formula, profile);
  }
  if (!profile || !wb.capabilities?.functionLocale) return formula;
  const ordinal = localeOrdinalForSpreadsheetProfile(profile);
  return canonicalizeFormula(wb, formula, ordinal);
}

/** Callback used by external clipboard payloads. Internal snapshots continue
 * to read formulas directly from the store and therefore remain canonical. */
export function localizeFormulaForClipboard(wb: WorkbookHandle): (formula: string) => string {
  if (typeof wb.localizeFormulaText === 'function') {
    return (formula) => wb.localizeFormulaText(formula);
  }
  if (typeof wb.getFormulaForDisplay === 'function') {
    return (formula) => wb.getFormulaForDisplay(formula);
  }
  return (formula) => formula;
}

/** Resolve the live workbook profile into the locale used by input coercion.
 * Older formulon builds do not expose the profile API, in which case the
 * parser intentionally keeps its en-US-compatible default. */
export function inputLocaleForWorkbook(wb: WorkbookHandle): string | undefined {
  try {
    const profile = wb.spreadsheetProfileId();
    if (!profile) return undefined;
    for (const [suffix, locale] of Object.entries(PROFILE_LOCALE_TO_INPUT_LOCALE)) {
      if (profile.endsWith(suffix)) return locale;
    }
  } catch {
    // A standalone/older workbook handle may not implement profile probing.
  }
  return undefined;
}

/**
 * Parse only the numeric grammar used by cell input, without applying any of
 * the other spreadsheet coercions. Dialog criteria use this boundary because
 * a date, boolean, formula, blank, or malformed numeric string is a textual
 * criterion rather than a number. The return value is always a finite number
 * so callers can safely pass it to native comparison APIs.
 */
export function parseLocalizedNumericInput(raw: string, locale?: string): number | null {
  const text = raw.trim();
  if (text === '' || text.startsWith('=') || text.startsWith("'")) return null;
  return parseNumberInput(text, locale)?.value ?? null;
}

/**
 * Map a user-typed string to the right primitive setter. Pure — depends on
 * neither the engine nor the store. Shared between the keyboard path,
 * formula bar, and clipboard paste so the rules stay in one place. Values,
 * and the number format a General cell picks up, follow desktop Excel.
 */
export function coerceInput(raw: string, options?: CoerceInputOptions): CoercedInput {
  const trimmed = raw.trim();
  if (trimmed === '') return { kind: 'blank' };
  if (trimmed.startsWith("'")) return { kind: 'text', value: trimmed.slice(1) };
  if (options?.forceText === true) return { kind: 'text', value: raw };
  if (trimmed.startsWith('=')) return { kind: 'formula', text: trimmed };
  const locale = options?.locale;
  // Logical and error literals match only without surrounding spaces.
  const [trueText, falseText] = booleanText(locale);
  // Full-width letters spell the same logical value.
  const upper = raw.normalize('NFKC').toUpperCase();
  if (upper === trueText) return { kind: 'bool', value: true };
  if (upper === falseText) return { kind: 'bool', value: false };
  const errorCode = errorLiteral(raw, locale);
  if (errorCode !== null) {
    return { kind: 'error', code: errorCode, text: formatError(errorCode, raw) };
  }
  const typed = parseTypedValue(trimmed, locale);
  if (typed === null) return { kind: 'text', value: raw };
  const implicitFormat = typed.formatCode ? formatCodeToNumFmt(typed.formatCode) : null;
  return implicitFormat
    ? { kind: 'number', value: typed.value, implicitFormat }
    : { kind: 'number', value: typed.value };
}

export function coerceInputForCell(
  state: State,
  a: Addr,
  raw: string,
  locale?: string,
  wb?: WorkbookHandle,
): CoercedInput {
  const format = wb ? effectiveCellFormatWithPending(wb, state, a) : formatWithPending(state, a);
  const coerced = coerceInput(raw, {
    forceText: format?.numFmt?.kind === 'text',
    locale,
  });
  if (wb && coerced.kind === 'formula') {
    return {
      ...coerced,
      text: canonicalizeLocalizedFormulaForInput(wb, coerced.text, locale),
    };
  }
  return coerced;
}

/**
 * Dispatch a coerced input to the workbook adapter. Callers can either use
 * this helper or call `coerceInput` and switch themselves (e.g. when paste
 * needs to skip writes for unchanged cells).
 */
export function writeCoerced(wb: WorkbookHandle, a: Addr, c: CoercedInput): void {
  switch (c.kind) {
    case 'blank':
      wb.setBlank(a);
      return;
    case 'formula':
      wb.setFormula(a, normalizeR1C1Formula(c.text, a));
      return;
    case 'number':
      wb.setNumber(a, c.value);
      return;
    case 'bool':
      wb.setBool(a, c.value);
      return;
    case 'error':
      wb.setError(a, c.code);
      return;
    case 'text':
      wb.setText(a, c.value);
  }
}

/** Attach a spreadsheet-style implicit number format (percent, currency, time,
 *  date) to a freshly-typed cell — but only while the cell is still on General
 *  format, so an explicit user format is never clobbered. */
function applyImplicitFormat(store: SpreadsheetStore, a: Addr, c: CoercedInput): void {
  if (c.kind !== 'number' || !c.implicitFormat) return;
  const current = formatWithPending(store.getState(), a)?.numFmt;
  if (current && current.kind !== 'general') return;
  mutators.setCellFormat(store, a, { numFmt: c.implicitFormat });
}

/** Convenience: coerce + write in one call for the common keyboard path.
 *  When `store` is supplied, sheet-protection is checked first — locked cells
 *  on protected sheets emit a console warning and the write is skipped. */
export function writeInput(
  wb: WorkbookHandle,
  a: Addr,
  raw: string,
  store?: SpreadsheetStore,
  locale?: string,
): void {
  const inputLocale = locale ?? inputLocaleForWorkbook(wb);
  let coerced: CoercedInput | null = null;
  if (store) {
    const state = store.getState();
    if (!isCellWritable(state, a)) {
      warnProtected(a);
      return;
    }
    coerced = coerceInputForCell(state, a, raw, inputLocale, wb);
  }
  const finalCoerced =
    coerced ??
    (() => {
      const input = coerceInput(raw, { locale: inputLocale });
      return input.kind === 'formula'
        ? { ...input, text: canonicalizeLocalizedFormulaForInput(wb, input.text, inputLocale) }
        : input;
    })();
  writeCoerced(wb, a, finalCoerced);
  if (store) applyImplicitFormat(store, a, finalCoerced);
}

/** Coerce + validate + write. When validation rejects with severity `stop`,
 *  the write is skipped and the outcome is returned so the caller can surface
 *  the error. Spreadsheet-compatible exception: rules with `showErrorMessage`
 *  disabled still record the invalid value, so Circle Invalid Data can flag it
 *  later. `warning` and `information` outcomes also write through but return
 *  the outcome for an inline toast. Range-backed list sources resolve against
 *  `wb` rooted at `a.sheet`. When `store` is supplied, sheet-protection is
 *  gated first; gated cells return `{ ok: true }` and emit a console warning
 *  rather than writing through. */
export function writeInputValidated(
  wb: WorkbookHandle,
  a: Addr,
  raw: string,
  validation: CellValidation | undefined,
  store?: SpreadsheetStore,
  locale?: string,
): ValidationOutcome {
  const inputLocale = locale ?? inputLocaleForWorkbook(wb);
  let coerced: CoercedInput | null = null;
  if (store) {
    const state = store.getState();
    if (!isCellWritable(state, a)) {
      warnProtected(a);
      return { ok: true };
    }
    coerced = coerceInputForCell(state, a, raw, inputLocale, wb);
  }
  if (!coerced) {
    const input = coerceInput(raw, { locale: inputLocale });
    coerced =
      input.kind === 'formula'
        ? { ...input, text: canonicalizeLocalizedFormulaForInput(wb, input.text, inputLocale) }
        : input;
  }
  if (!validation) {
    writeCoerced(wb, a, coerced);
    if (store) applyImplicitFormat(store, a, coerced);
    return { ok: true };
  }
  const outcome = validateCoercedInput(wb, a, coerced, validation);
  if (outcome.ok || outcome.severity !== 'stop') {
    writeCoerced(wb, a, coerced);
    if (store) applyImplicitFormat(store, a, coerced);
  }
  return outcome;
}

/** Validate `coerced` for `a` without writing. Only a `stop` outcome blocks
 *  the entry; custom rules evaluate against `wb`. */
export function validateCoercedInput(
  wb: WorkbookHandle,
  a: Addr,
  coerced: CoercedInput,
  validation: CellValidation,
): ValidationOutcome {
  const nativeOutcome = nativeValidationOutcome(wb, a, coerced, validation);
  if (nativeOutcome) {
    const silentStop =
      !nativeOutcome.ok &&
      nativeOutcome.severity === 'stop' &&
      validation.showErrorMessage === false;
    return silentStop ? { ok: true } : nativeOutcome;
  }
  const evalCustom = validation.kind === 'custom' ? makeCustomEvaluator(wb, a, coerced) : undefined;
  const outcome = validateAgainst(validation, coerced, makeRangeResolver(wb, a.sheet), evalCustom);
  // A `stop` rule with the error alert disabled records the invalid value
  // silently — no blocking dialog — so it must report success to the caller
  // rather than a `stop` rejection that keeps the editor open.
  const silentStop =
    !outcome.ok && outcome.severity === 'stop' && validation.showErrorMessage === false;
  return silentStop ? { ok: true } : outcome;
}

/** Build a custom-validation evaluator. Excel's custom rule is a formula that
 *  must evaluate to TRUE for the entry to be accepted. The engine only exposes
 *  a fresh-workbook formula evaluator, so we substitute each single-cell
 *  reference with a literal — the *candidate* value for the cell under test,
 *  and current values for any other referenced cells — then evaluate the
 *  resulting self-contained expression. Ranges and cross-sheet references
 *  cannot be substituted this way, so those rules return `null` (accept for
 *  parity) rather than block. */
function makeCustomEvaluator(
  wb: WorkbookHandle,
  a: Addr,
  candidate: CoercedInput,
): CustomValidationEvaluator {
  return (formula: string): boolean | null => {
    const literal = substituteCellRefs(formula, wb, a, candidate);
    if (literal === null) return null;
    const res = wb.evalFormula(literal);
    if (res.status.status !== 0) return null; // evaluation unavailable → parity accept
    const v = fromEngineValue(res.value);
    if (v.kind === 'number') return v.value !== 0;
    if (v.kind === 'bool') return v.value;
    return false; // text / error / blank → not satisfied (Excel requires TRUE / nonzero)
  };
}

/** Render a coerced input as a formula literal token. */
function coercedLiteral(input: CoercedInput): string {
  switch (input.kind) {
    case 'number':
      return String(input.value);
    case 'bool':
      return input.value ? 'TRUE' : 'FALSE';
    case 'error':
      return formatError(input.code, input.text);
    case 'text':
      return JSON.stringify(input.value);
    case 'formula':
      return input.text.startsWith('=') ? input.text.slice(1) : input.text;
    default:
      return '0'; // blank → 0 in a numeric context
  }
}

/** Render an engine cell value as a formula literal token. */
function cellValueLiteral(v: CellValue): string {
  switch (v.kind) {
    case 'number':
      return String(v.value);
    case 'bool':
      return v.value ? 'TRUE' : 'FALSE';
    case 'text':
      return JSON.stringify(v.value);
    default:
      return '0';
  }
}

/** Substitute every single-cell, same-sheet reference in `formula` with a
 *  literal value. Returns null when the formula contains a range or a
 *  sheet-qualified reference (unsupported for literal substitution). */
function substituteCellRefs(
  formula: string,
  wb: WorkbookHandle,
  a: Addr,
  candidate: CoercedInput,
): string | null {
  const refs = extractRefs(formula);
  if (refs.length === 0) return formula;
  const replacements: { start: number; end: number; text: string }[] = [];
  for (const ref of refs) {
    // Only single cells are substitutable; a range would collapse incorrectly.
    if (ref.r0 !== ref.r1 || ref.c0 !== ref.c1) return null;
    // Sheet-qualified refs (contain `!`) can't be resolved here.
    if (formula.slice(ref.start, ref.end).includes('!')) return null;
    const isTarget = ref.r0 === a.row && ref.c0 === a.col;
    const text = isTarget
      ? coercedLiteral(candidate)
      : cellValueLiteral(wb.getValue({ sheet: a.sheet, row: ref.r0, col: ref.c0 }));
    replacements.push({ start: ref.start, end: ref.end, text });
  }
  let out = formula;
  for (const rep of replacements.sort((x, y) => y.start - x.start)) {
    out = `${out.slice(0, rep.start)}${rep.text}${out.slice(rep.end)}`;
  }
  return out;
}
