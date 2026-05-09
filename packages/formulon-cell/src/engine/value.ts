import { type CellValue, type Value, ValueKind } from './types.js';

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
};

const BLANK: CellValue = { kind: 'blank' };

export function formatGeneralNumber(value: number, locale = 'en-US'): string {
  if (!Number.isFinite(value)) return String(value);
  const abs = Math.abs(value);
  if (abs > 0 && (abs >= 1e11 || abs < 1e-9)) {
    return value.toExponential(5).replace(/e([+-]?)(\d+)$/i, (_m, sign: string, exp: string) => {
      const normalizedSign = sign === '-' ? '-' : '+';
      return `E${normalizedSign}${exp.padStart(2, '0')}`;
    });
  }
  return new Intl.NumberFormat(locale, { maximumFractionDigits: 12 }).format(value);
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
      return { kind: 'error', code: v.errorCode, text: ERROR_NAME[v.errorCode] ?? '#ERR!' };
    default:
      return BLANK;
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
      return v.value ? 'TRUE' : 'FALSE';
    case 'text':
      return v.value;
    case 'error':
      return v.text;
    default:
      return '';
  }
}
