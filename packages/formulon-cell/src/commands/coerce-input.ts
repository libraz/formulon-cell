import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellValidation, SpreadsheetStore } from '../store/store.js';
import { isCellWritable, warnProtected } from './protection.js';
import { type ValidationOutcome, validateAgainst } from './validate.js';

export type CoercedInput =
  | { kind: 'blank' }
  | { kind: 'formula'; text: string }
  | { kind: 'number'; value: number }
  | { kind: 'bool'; value: boolean }
  | { kind: 'text'; value: string };

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

const parseNumericValue = (raw: string): number | null => {
  let text = raw.trim();
  let negative = false;
  if (text.startsWith('(') && text.endsWith(')')) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  text = text.replace(CURRENCY, '');
  const isPercent = text.endsWith('%');
  if (isPercent) text = text.slice(0, -1);
  if (!NUMERIC.test(text)) return null;
  const n = Number(text.replaceAll(',', ''));
  if (Number.isNaN(n)) return null;
  const signed = negative ? -Math.abs(n) : n;
  return isPercent ? signed / 100 : signed;
};

const parseTimeValue = (raw: string): number | null => {
  const m = TIME.exec(raw.trim());
  if (!m) return null;
  const sign = m[1] === '-' ? -1 : 1;
  let hours = Number(m[2]);
  const minutes = Number(m[3]);
  const seconds = m[4] === undefined ? 0 : Number(m[4]);
  const meridiem = m[5]?.toUpperCase();
  if (meridiem) {
    if (hours < 1 || hours > 12) return null;
    if (meridiem === 'AM') hours = hours === 12 ? 0 : hours;
    else hours = hours === 12 ? 12 : hours + 12;
  }
  return (sign * (hours * 3600 + minutes * 60 + seconds)) / 86_400;
};

/**
 * Map a user-typed string to the right primitive setter. Pure — depends on
 * neither the engine nor the store. Shared between the keyboard path,
 * formula bar, and clipboard paste so the rules stay in one place.
 */
export function coerceInput(raw: string): CoercedInput {
  const trimmed = raw.trim();
  const numericTrimmed = normalizeNumericText(trimmed);
  if (trimmed === '') return { kind: 'blank' };
  if (trimmed.startsWith('=')) return { kind: 'formula', text: trimmed };
  if (trimmed.startsWith("'")) return { kind: 'text', value: trimmed.slice(1) };
  const boolText = numericTrimmed.toUpperCase();
  if (boolText === 'TRUE' || boolText === 'FALSE') {
    return { kind: 'bool', value: boolText === 'TRUE' };
  }
  const time = parseTimeValue(numericTrimmed);
  if (time !== null) return { kind: 'number', value: time };
  const n = parseNumericValue(numericTrimmed);
  if (n !== null) return { kind: 'number', value: n };
  return { kind: 'text', value: raw };
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
      wb.setFormula(a, c.text);
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

/** Convenience: coerce + write in one call for the common keyboard path.
 *  When `store` is supplied, sheet-protection is checked first — locked cells
 *  on protected sheets emit a console warning and the write is skipped. */
export function writeInput(
  wb: WorkbookHandle,
  a: Addr,
  raw: string,
  store?: SpreadsheetStore,
): void {
  if (store && !isCellWritable(store.getState(), a)) {
    warnProtected(a);
    return;
  }
  writeCoerced(wb, a, coerceInput(raw));
}

/** Coerce + validate + write. When validation rejects with severity `stop`,
 *  the write is skipped and the outcome is returned so the caller can surface
 *  the error. `warning` and `information` outcomes still write through but
 *  the message is returned for an inline toast. Range-backed list sources
 *  resolve against `wb` rooted at `a.sheet`. When `store` is supplied,
 *  sheet-protection is gated first; gated cells return `{ ok: true }` and
 *  emit a console warning rather than writing through. */
export function writeInputValidated(
  wb: WorkbookHandle,
  a: Addr,
  raw: string,
  validation: CellValidation | undefined,
  store?: SpreadsheetStore,
): ValidationOutcome {
  if (store && !isCellWritable(store.getState(), a)) {
    warnProtected(a);
    return { ok: true };
  }
  const coerced = coerceInput(raw);
  if (!validation) {
    writeCoerced(wb, a, coerced);
    return { ok: true };
  }
  const outcome = validateAgainst(validation, coerced, makeRangeResolver(wb, a.sheet));
  if (outcome.ok || outcome.severity !== 'stop') {
    writeCoerced(wb, a, coerced);
  }
  return outcome;
}
