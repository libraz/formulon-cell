import { intlFormatterCache } from '../format/intl-cache.js';
import type { Addr } from './types.js';
import { type CellValue, type Value, ValueKind } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const ERROR_NAME: Readonly<Record<number, string>> = {
  0: '#NULL!',
  1: '#DIV/0!',
  2: '#VALUE!',
  3: '#REF!',
  4: '#NAME?',
  5: '#NUM!',
  6: '#N/A',
  7: '#GETTING_DATA',
  8: '#SPILL!',
  9: '#CALC!',
  10: '#FIELD!',
  11: '#BLOCKED!',
  12: '#CONNECT!',
  13: '#EXTERNAL!',
  14: '#BUSY!',
  15: '#PYTHON!',
  16: '#UNKNOWN!',
};

const BLANK: CellValue = { kind: 'blank' };

const localeLanguage = (locale: string): string =>
  locale.trim().toLowerCase().replaceAll('_', '-').split('-')[0] ?? 'en';

/**
 * Render Excel's scientific notation using the active locale's decimal mark.
 * `Number#toExponential` remains the source of truth for rounding and digit
 * selection; Intl is used only to discover the locale decimal separator.
 */
export function formatScientificNumber(
  value: number,
  decimals: number,
  locale = 'en-US',
  exponentDigits = 2,
): string {
  if (!Number.isFinite(value)) return String(value);
  const precision = Number.isFinite(decimals) ? Math.max(0, Math.trunc(decimals)) : 0;
  const exponentWidth = Number.isFinite(exponentDigits)
    ? Math.max(0, Math.trunc(exponentDigits))
    : 2;
  const decimal =
    intlFormatterCache
      .numberFormat(locale, {
        useGrouping: false,
        minimumFractionDigits: 1,
        maximumFractionDigits: 1,
      })
      .formatToParts(1.1)
      .find((part) => part.type === 'decimal')?.value ?? '.';
  const scientific = value.toExponential(precision);
  const match = /^([+-]?\d(?:\.\d+)?)[eE]([+-]?)(\d+)$/.exec(scientific);
  if (!match) return scientific;
  const mantissa = (match[1] ?? '').replace('.', decimal);
  const sign = match[2] === '-' ? '-' : '+';
  const exponent = (match[3] ?? '').padStart(exponentWidth, '0');
  return `${mantissa}E${sign}${exponent}`;
}

/** Render an engine error using the measured locale token where available. */
export function formatError(code: number, fallbackText: string, locale = 'en-US'): string {
  const canonical = ERROR_NAME[code];
  if (canonical === undefined) return fallbackText;
  // formulon currently has a measured German spelling only for #N/A. Keep all
  // other error names canonical until their locale captures exist.
  if (code === 6 && localeLanguage(locale) === 'de') return '#NV';
  return canonical;
}

/** Render a logical value using the display token of the active UI profile. */
export function formatBoolean(value: boolean, locale = 'en-US'): string {
  const tokens: readonly [string, string] =
    localeLanguage(locale) === 'de'
      ? ['WAHR', 'FALSCH']
      : localeLanguage(locale) === 'fr'
        ? ['VRAI', 'FAUX']
        : ['TRUE', 'FALSE'];
  return value ? tokens[0] : tokens[1];
}

/** Render a number as edit input: preserve String(number)'s precision and
 *  change only the decimal mark used by German/French input profiles. */
export function formatEditableNumber(value: number, locale = 'en-US'): string {
  const text = String(value);
  const language = localeLanguage(locale);
  return language === 'de' || language === 'fr' ? text.replace('.', ',') : text;
}

export function formatGeneralNumber(
  value: number,
  locale = 'en-US',
  options: { useGrouping?: boolean } = {},
): string {
  if (!Number.isFinite(value)) return String(value);
  const abs = Math.abs(value);
  if (abs > 0 && (abs >= 1e11 || abs < 1e-9)) {
    return formatScientificNumber(value, 5, locale);
  }
  return intlFormatterCache
    .numberFormat(locale, {
      maximumFractionDigits: 12,
      ...options,
    })
    .format(value);
}

export function fromEngineValue(v: Value): CellValue {
  switch (v.kind) {
    case ValueKind.Number:
      return { kind: 'number', value: v.number };
    case ValueKind.Bool:
      return { kind: 'bool', value: v.boolean !== 0 };
    case ValueKind.Text:
      return { kind: 'text', value: v.text };
    case ValueKind.Error:
      return { kind: 'error', code: v.errorCode, text: formatError(v.errorCode, '#ERR!') };
    default:
      return BLANK;
  }
}

/** Write a stored cell through the typed `WorkbookHandle` API: the formula
 *  when present, otherwise the value (blank for anything else). */
export function writeCell(
  wb: WorkbookHandle,
  addr: Addr,
  value: CellValue,
  formula: string | null,
): void {
  if (formula) {
    wb.setFormula(addr, formula);
    return;
  }
  switch (value.kind) {
    case 'number':
      wb.setNumber(addr, value.value);
      return;
    case 'text':
      wb.setText(addr, value.value);
      return;
    case 'bool':
      wb.setBool(addr, value.value);
      return;
    case 'error':
      wb.setError(addr, value.code);
      return;
    default:
      wb.setBlank(addr);
  }
}

/** Format a CellValue for display. Numbers honour locale; nothing fancier yet
 *  (number formats arrive in 1.0 with the engine `setNumberFormat` API). */
export function formatCell(v: CellValue, locale = 'en-US'): string {
  switch (v.kind) {
    case 'blank':
      return '';
    case 'number':
      return formatGeneralNumber(v.value, locale);
    case 'bool':
      return formatBoolean(v.value, locale);
    case 'text':
      return v.value;
    case 'error':
      return formatError(v.code, v.text, locale);
    default:
      return '';
  }
}
