import { effectiveCellFormatWithPending } from '../engine/cell-format-sync.js';
import {
  canonicalizeFormula,
  localeOrdinalForSpreadsheetProfile,
} from '../engine/function-locale.js';
import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Addr, CellValue, SpreadsheetProfileId } from '../engine/types.js';
import { fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { excelSerialFromParts } from '../format/date-serial.js';
import { spreadsheetProfileForLocale } from '../mount/locale-profile.js';
import { formatWithPending } from '../store/pending-format.js';
import {
  type CellValidation,
  mutators,
  type NumFmt,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
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
  | { kind: 'text'; value: string };

export interface CoerceInputOptions {
  forceText?: boolean;
  /** BCP47 input locale. When omitted, the input grammar uses en-US. */
  locale?: string;
}

const CURRENCY_SYMBOL = /[¥$€£₩฿]/;
const CURRENCY_PREFIX = /^([¥$€£₩฿])\s*/;
const CURRENCY_SUFFIX = /([¥$€£₩฿])$/;
const TIME = /^([+-]?)(\d+):([0-5]\d)(?::([0-5]\d))?(?:\s*([AP]M))?$/i;

type InputDateOrder = 'mdy' | 'dmy' | 'ymd';

interface InputLocaleRules {
  readonly decimal: string;
  readonly groupSeparators: readonly string[];
  readonly dateOrder: InputDateOrder;
  readonly dateSeparator: string;
}

const INPUT_LOCALE_RULES: Readonly<Record<string, InputLocaleRules>> = {
  en: { decimal: '.', groupSeparators: [','], dateOrder: 'mdy', dateSeparator: '/' },
  ja: { decimal: '.', groupSeparators: [','], dateOrder: 'ymd', dateSeparator: '/' },
  zh: { decimal: '.', groupSeparators: [','], dateOrder: 'ymd', dateSeparator: '/' },
  ko: { decimal: '.', groupSeparators: [','], dateOrder: 'ymd', dateSeparator: '/' },
  th: { decimal: '.', groupSeparators: [','], dateOrder: 'dmy', dateSeparator: '/' },
  de: { decimal: ',', groupSeparators: ['.'], dateOrder: 'dmy', dateSeparator: '.' },
  fr: {
    decimal: ',',
    groupSeparators: ['\u202f', '\u00a0', ' '],
    dateOrder: 'dmy',
    dateSeparator: '/',
  },
};

const EN_INPUT_LOCALE_RULES: InputLocaleRules = {
  decimal: '.',
  groupSeparators: [','],
  dateOrder: 'mdy',
  dateSeparator: '/',
};

const inputLanguage = (locale: string | undefined): string =>
  (locale?.replace('_', '-').split('-', 1)[0] ?? 'en').toLowerCase();

const inputLocaleRules = (locale: string | undefined): InputLocaleRules =>
  INPUT_LOCALE_RULES[inputLanguage(locale)] ?? EN_INPUT_LOCALE_RULES;

// The engine's seven profile definitions use TRUE/FALSE for en, ja, zh, ko,
// and th. German and French are the only profiles with localized boolean text.
// Keep these spellings scoped to their matching language; TRUE/FALSE remains
// accepted for every locale for formula and workbook compatibility.
const LOCALIZED_BOOLEAN_TEXT: Readonly<Record<string, readonly [string, string]>> = {
  de: ['WAHR', 'FALSCH'],
  fr: ['VRAI', 'FAUX'],
};

const PROFILE_LOCALE_TO_INPUT_LOCALE: Readonly<Record<string, string>> = {
  '-ja_JP': 'ja-JP',
  '-en_US': 'en-US',
  '-de_DE': 'de-DE',
  '-fr_FR': 'fr-FR',
  '-zh_CN': 'zh-CN',
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

const normalizeNumericText = (raw: string): string =>
  [...raw]
    .map((ch) => {
      const cp = ch.codePointAt(0) ?? 0;
      if (cp >= 0xff01 && cp <= 0xff5e) return String.fromCodePoint(cp - 0xfee0);
      if (cp === 0xffe5) return '¥';
      if (cp === 0x3000) return ' ';
      return ch;
    })
    .join('');

/** A parsed numeric input plus the number format spreadsheets implicitly
 *  attach to it (percent, currency, time, date). `format` is left undefined for
 *  a bare number so the cell keeps its General format. */
interface ParsedNumber {
  value: number;
  format?: NumFmt;
}

interface ParsedLocalizedNumber {
  readonly canonical: string;
  readonly decimals: number;
}

const DIGITS = /^\d+$/;

/** Parse a locale's decimal/grouping grammar without accepting malformed
 * grouping by simply deleting punctuation. For example, `1.23,5` is text in
 * de-DE because the period would have to separate three-digit groups. */
const parseLocalizedNumber = (
  raw: string,
  rules: InputLocaleRules,
): ParsedLocalizedNumber | null => {
  const exponentMatch = /^(.+?)([eE][+-]?\d+)$/.exec(raw);
  const mantissa = exponentMatch?.[1] ?? raw;
  const exponent = exponentMatch?.[2] ?? '';
  let sign = '';
  let unsigned = mantissa;
  if (unsigned.startsWith('+') || unsigned.startsWith('-')) {
    sign = unsigned[0] ?? '';
    unsigned = unsigned.slice(1);
  }
  if (unsigned === '') return null;

  const decimalIndex = unsigned.indexOf(rules.decimal);
  if (decimalIndex !== unsigned.lastIndexOf(rules.decimal)) return null;
  const integerPart = decimalIndex < 0 ? unsigned : unsigned.slice(0, decimalIndex);
  const fractionPart = decimalIndex < 0 ? '' : unsigned.slice(decimalIndex + 1);
  if (fractionPart !== '' && !DIGITS.test(fractionPart)) return null;

  let integerDigits = integerPart;
  if (integerPart === '') {
    if (fractionPart === '') return null;
    integerDigits = '0';
  } else if (!DIGITS.test(integerPart)) {
    const presentSeparators = rules.groupSeparators.filter((separator) =>
      integerPart.includes(separator),
    );
    // French accepts three kinds of space as a group separator, but a single
    // number must use one kind consistently.
    if (presentSeparators.length !== 1) return null;
    const separator = presentSeparators[0] ?? '';
    const groups = integerPart.split(separator);
    if (
      groups.length < 2 ||
      !DIGITS.test(groups[0] ?? '') ||
      (groups[0]?.length ?? 0) > 3 ||
      groups.slice(1).some((group) => group.length !== 3 || !DIGITS.test(group))
    ) {
      return null;
    }
    integerDigits = groups.join('');
  }

  return {
    canonical: `${sign}${integerDigits}${decimalIndex < 0 ? '' : `.${fractionPart}`}${exponent}`,
    decimals: decimalIndex < 0 ? 0 : fractionPart.length,
  };
};

const parseNumericValue = (raw: string, locale?: string): ParsedNumber | null => {
  let text = raw.trim();
  let negative = false;
  if (text.startsWith('(') && text.endsWith(')')) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  const rules = inputLocaleRules(locale);
  const language = inputLanguage(locale);
  let symbol: string | null = null;
  const prefix = CURRENCY_PREFIX.exec(text);
  if (prefix) {
    symbol = prefix[1] ?? null;
    text = text.slice(prefix[0].length);
  }
  const suffix = CURRENCY_SUFFIX.exec(text);
  if (suffix) {
    // Native German/French currency formatting commonly places the symbol
    // after the number, separated by a regular, NBSP, or narrow NBSP space.
    // Keep the accepted suffix locale-scoped so an English trailing symbol
    // remains ordinary text, as it did before locale-aware parsing.
    if (language !== 'de' && language !== 'fr') return null;
    if (symbol !== null) return null;
    symbol = suffix[1] ?? null;
    text = text.slice(0, suffix.index).trimEnd();
  }
  if (CURRENCY_SYMBOL.test(text)) return null;
  const isPercent = text.endsWith('%');
  if (isPercent) text = text.slice(0, -1);
  text = text.trim();
  const parsed = parseLocalizedNumber(text, rules);
  if (!parsed) return null;
  const n = Number(parsed.canonical);
  if (!Number.isFinite(n)) return null;
  const signed = negative ? -Math.abs(n) : n;
  if (isPercent) {
    return {
      value: signed / 100,
      format: { kind: 'percent', decimals: parsed.decimals },
    };
  }
  if (symbol) {
    return { value: signed, format: { kind: 'currency', decimals: 2, symbol } };
  }
  return { value: signed };
};

/**
 * Parse only the numeric grammar used by cell input, without applying any of
 * the other spreadsheet coercions. Dialog criteria use this boundary because
 * a date, boolean, formula, blank, or malformed numeric string is a textual
 * criterion rather than a number. The return value is always a finite number
 * so callers can safely pass it to native comparison APIs.
 */
export function parseLocalizedNumericInput(raw: string, locale?: string): number | null {
  const text = normalizeNumericText(raw.trim());
  if (text === '' || text.startsWith('=') || text.startsWith("'")) return null;
  const parsed = parseNumericValue(text, locale);
  return parsed && Number.isFinite(parsed.value) ? parsed.value : null;
}

const parseTimeValue = (raw: string): ParsedNumber | null => {
  const m = TIME.exec(raw.trim());
  if (!m) return null;
  const sign = m[1] === '-' ? -1 : 1;
  let hours = Number(m[2]);
  const minutes = Number(m[3]);
  const hasSeconds = m[4] !== undefined;
  const seconds = hasSeconds ? Number(m[4]) : 0;
  const meridiem = m[5]?.toUpperCase();
  if (meridiem) {
    if (hours < 1 || hours > 12) return null;
    if (meridiem === 'AM') hours = hours === 12 ? 0 : hours;
    else hours = hours === 12 ? 12 : hours + 12;
  }
  const value = (sign * (hours * 3600 + minutes * 60 + seconds)) / 86_400;
  const elapsed = !meridiem && hours >= 24;
  const pattern = elapsed
    ? hasSeconds
      ? '[h]:mm:ss'
      : '[h]:mm'
    : meridiem
      ? hasSeconds
        ? 'h:mm:ss AM/PM'
        : 'h:mm AM/PM'
      : hasSeconds
        ? 'h:mm:ss'
        : 'h:mm';
  return { value, format: { kind: 'time', pattern } };
};

const MONTH_NAMES: Record<string, number> = {
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
};

/** Excel two-digit-year window: 00–29 → 2000s, 30–99 → 1900s. */
const normalizeYear = (y: number): number => (y < 100 ? (y < 30 ? 2000 + y : 1900 + y) : y);

/** Convert a Y/M/D triple to a spreadsheet serial, validating the calendar
 *  date. Returns null for impossible dates (e.g. 13/45) so the caller falls
 *  back to text. */
const ymdToSerial = (year: number, month: number, day: number): number | null => {
  return excelSerialFromParts(year, month, day);
};

const parseDateValue = (raw: string, locale?: string): ParsedNumber | null => {
  const text = raw.trim();
  // ISO / dash-first: 2024-12-25 or 2024/12/25.
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text);
  if (m) {
    const serial = ymdToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
    return serial === null
      ? null
      : { value: serial, format: { kind: 'date', pattern: 'yyyy-mm-dd' } };
  }

  const rules = inputLocaleRules(locale);
  const numericPattern =
    rules.dateOrder === 'ymd'
      ? /^(\d{2,4})[-/](\d{1,2})[-/](\d{1,2})$/
      : rules.dateOrder === 'mdy'
        ? /^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/
        : rules.dateOrder === 'dmy' && rules.dateSeparator === '.'
          ? /^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?$/
          : new RegExp(
              `^(\\d{1,2})${rules.dateSeparator}(\\d{1,2})(?:${rules.dateSeparator}(\\d{2,4}))?$`,
            );
  m = numericPattern.exec(text);
  if (m) {
    const first = Number(m[1]);
    const second = Number(m[2]);
    const yearGiven = m[3] !== undefined;
    if (rules.dateOrder === 'ymd') {
      const year = normalizeYear(first);
      const serial = ymdToSerial(year, second, Number(m[3]));
      if (serial === null) return null;
      return { value: serial, format: { kind: 'date', pattern: 'yyyy/m/d' } };
    }
    const month = rules.dateOrder === 'mdy' ? first : second;
    const day = rules.dateOrder === 'mdy' ? second : first;
    const year = yearGiven ? normalizeYear(Number(m[3])) : new Date().getUTCFullYear();
    const pattern =
      rules.dateOrder === 'mdy'
        ? 'm/d/yyyy'
        : rules.dateSeparator === '.'
          ? 'd.m.yyyy'
          : 'd/m/yyyy';
    const serial = ymdToSerial(year, month, day);
    if (serial === null) return null;
    return {
      value: serial,
      format: { kind: 'date', pattern: yearGiven ? pattern : pattern.slice(0, -5) },
    };
  }
  // Textual month: 25-Dec-2024, 25 Dec 2024, Dec 25 2024, Dec 25, 2024.
  m = /^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{2,4})$/.exec(text);
  if (m) {
    const month = MONTH_NAMES[(m[2] ?? '').slice(0, 3).toLowerCase()];
    if (month === undefined) return null;
    const serial = ymdToSerial(normalizeYear(Number(m[3])), month, Number(m[1]));
    return serial === null
      ? null
      : { value: serial, format: { kind: 'date', pattern: 'd-mmm-yyyy' } };
  }
  m = /^([A-Za-z]{3,})[-\s](\d{1,2}),?[-\s](\d{2,4})$/.exec(text);
  if (m) {
    const month = MONTH_NAMES[(m[1] ?? '').slice(0, 3).toLowerCase()];
    if (month === undefined) return null;
    const serial = ymdToSerial(normalizeYear(Number(m[3])), month, Number(m[2]));
    return serial === null
      ? null
      : { value: serial, format: { kind: 'date', pattern: 'd-mmm-yyyy' } };
  }
  return null;
};

/**
 * Map a user-typed string to the right primitive setter. Pure — depends on
 * neither the engine nor the store. Shared between the keyboard path,
 * formula bar, and clipboard paste so the rules stay in one place.
 */
export function coerceInput(raw: string, options?: CoerceInputOptions): CoercedInput {
  const trimmed = raw.trim();
  const numericTrimmed = normalizeNumericText(trimmed);
  if (trimmed === '') return { kind: 'blank' };
  if (trimmed.startsWith("'")) return { kind: 'text', value: trimmed.slice(1) };
  if (options?.forceText === true) return { kind: 'text', value: raw };
  if (trimmed.startsWith('=')) return { kind: 'formula', text: trimmed };
  const boolText = trimmed.toUpperCase();
  if (boolText === 'TRUE' || boolText === 'FALSE') {
    return { kind: 'bool', value: boolText === 'TRUE' };
  }
  const localizedBooleanText = LOCALIZED_BOOLEAN_TEXT[inputLanguage(options?.locale)];
  if (localizedBooleanText) {
    if (boolText === localizedBooleanText[0]) return { kind: 'bool', value: true };
    if (boolText === localizedBooleanText[1]) return { kind: 'bool', value: false };
  }
  const time = parseTimeValue(numericTrimmed);
  if (time !== null) return numberInput(time);
  const n = parseNumericValue(numericTrimmed, options?.locale);
  if (n !== null) return numberInput(n);
  const date = parseDateValue(numericTrimmed, options?.locale);
  if (date !== null) return numberInput(date);
  return { kind: 'text', value: raw };
}

/** Wrap a parsed number, attaching its implicit format only when present. */
function numberInput(parsed: ParsedNumber): CoercedInput {
  return parsed.format
    ? { kind: 'number', value: parsed.value, implicitFormat: parsed.format }
    : { kind: 'number', value: parsed.value };
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
