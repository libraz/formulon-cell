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
    case ValueKind.Blank:
    case ValueKind.Array:
    case ValueKind.Ref:
    case ValueKind.Lambda:
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
      return Number.isFinite(v.value)
        ? new Intl.NumberFormat(locale, { maximumFractionDigits: 12 }).format(v.value)
        : String(v.value);
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
