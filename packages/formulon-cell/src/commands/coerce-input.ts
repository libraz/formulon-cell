import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Addr, CellValue } from '../engine/types.js';
import { fromEngineValue } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
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
}

const NUMERIC = /^[+-]?(?:(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;
const CURRENCY = /^[¥$€£]\s*/;
const TIME = /^([+-]?)(\d+):([0-5]\d)(?::([0-5]\d))?(?:\s*([AP]M))?$/i;

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

const countDecimals = (digits: string): number => {
  const dot = digits.indexOf('.');
  return dot < 0 ? 0 : digits.length - dot - 1;
};

const parseNumericValue = (raw: string): ParsedNumber | null => {
  let text = raw.trim();
  let negative = false;
  if (text.startsWith('(') && text.endsWith(')')) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  const symbol = CURRENCY.exec(text)?.[0]?.trim() ?? null;
  text = text.replace(CURRENCY, '');
  const isPercent = text.endsWith('%');
  if (isPercent) text = text.slice(0, -1);
  text = text.trim();
  if (!NUMERIC.test(text)) return null;
  const digits = text.replaceAll(',', '');
  const n = Number(digits);
  if (Number.isNaN(n)) return null;
  const signed = negative ? -Math.abs(n) : n;
  if (isPercent) {
    return { value: signed / 100, format: { kind: 'percent', decimals: countDecimals(digits) } };
  }
  if (symbol) {
    return { value: signed, format: { kind: 'currency', decimals: 2, symbol } };
  }
  return { value: signed };
};

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
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const ms = Date.UTC(year, month - 1, day);
  const d = new Date(ms);
  // Reject overflow (e.g. Feb 30 rolls into March).
  if (d.getUTCMonth() !== month - 1 || d.getUTCDate() !== day) return null;
  return ms / 86_400_000 + 25569;
};

const parseDateValue = (raw: string): ParsedNumber | null => {
  const text = raw.trim();
  // ISO / dash-first: 2024-12-25 or 2024/12/25.
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text);
  if (m) {
    const serial = ymdToSerial(Number(m[1]), Number(m[2]), Number(m[3]));
    return serial === null
      ? null
      : { value: serial, format: { kind: 'date', pattern: 'yyyy-mm-dd' } };
  }
  // US numeric: 12/25/2024, 12-25-2024, or 12/25 (current year).
  m = /^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/.exec(text);
  if (m) {
    const month = Number(m[1]);
    const day = Number(m[2]);
    const yearGiven = m[3] !== undefined;
    const year = yearGiven ? normalizeYear(Number(m[3])) : new Date().getUTCFullYear();
    const serial = ymdToSerial(year, month, day);
    if (serial === null) return null;
    return { value: serial, format: { kind: 'date', pattern: yearGiven ? 'm/d/yyyy' : 'm/d' } };
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
  const time = parseTimeValue(numericTrimmed);
  if (time !== null) return numberInput(time);
  const n = parseNumericValue(numericTrimmed);
  if (n !== null) return numberInput(n);
  const date = parseDateValue(numericTrimmed);
  if (date !== null) return numberInput(date);
  return { kind: 'text', value: raw };
}

/** Wrap a parsed number, attaching its implicit format only when present. */
function numberInput(parsed: ParsedNumber): CoercedInput {
  return parsed.format
    ? { kind: 'number', value: parsed.value, implicitFormat: parsed.format }
    : { kind: 'number', value: parsed.value };
}

export function coerceInputForCell(state: State, a: Addr, raw: string): CoercedInput {
  const format = formatWithPending(state, a);
  return coerceInput(raw, {
    forceText: format?.numFmt?.kind === 'text',
  });
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
): void {
  let coerced: CoercedInput | null = null;
  if (store) {
    const state = store.getState();
    if (!isCellWritable(state, a)) {
      warnProtected(a);
      return;
    }
    coerced = coerceInputForCell(state, a, raw);
  }
  const finalCoerced = coerced ?? coerceInput(raw);
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
): ValidationOutcome {
  let coerced: CoercedInput | null = null;
  if (store) {
    const state = store.getState();
    if (!isCellWritable(state, a)) {
      warnProtected(a);
      return { ok: true };
    }
    coerced = coerceInputForCell(state, a, raw);
  }
  coerced ??= coerceInput(raw);
  if (!validation) {
    writeCoerced(wb, a, coerced);
    if (store) applyImplicitFormat(store, a, coerced);
    return { ok: true };
  }
  const evalCustom = validation.kind === 'custom' ? makeCustomEvaluator(wb, a, coerced) : undefined;
  const outcome = validateAgainst(validation, coerced, makeRangeResolver(wb, a.sheet), evalCustom);
  // A `stop` rule with the error alert disabled records the invalid value
  // silently — no blocking dialog — so it must report success to the caller
  // rather than a `stop` rejection that keeps the editor open.
  const silentStop =
    !outcome.ok && outcome.severity === 'stop' && validation.showErrorMessage === false;
  if (outcome.ok || outcome.severity !== 'stop' || silentStop) {
    writeCoerced(wb, a, coerced);
    if (store) applyImplicitFormat(store, a, coerced);
  }
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
