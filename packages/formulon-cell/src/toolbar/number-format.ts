import { normalizeFormatLocale } from '../format/locale.js';
import type { NumFmt } from '../store/types.js';

export type NumberFormatAction =
  | 'general'
  | 'fixed'
  | 'currency'
  | 'accounting'
  | 'shortDate'
  | 'longDate'
  | 'time'
  | 'percent'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'more';

const defaultCurrencySymbolForToolbar = (locale: string): string =>
  normalizeFormatLocale(locale).startsWith('ja') ? '¥' : '$';

export const numberFormatForAction = (
  action: NumberFormatAction,
  locale: string,
): NumFmt | null => {
  const isJapanese = normalizeFormatLocale(locale).startsWith('ja');
  const symbol = defaultCurrencySymbolForToolbar(locale);
  switch (action) {
    case 'general':
      return { kind: 'general' };
    case 'fixed':
      return { kind: 'fixed', decimals: 0 };
    case 'currency':
      return { kind: 'currency', decimals: 2, symbol };
    case 'accounting':
      return { kind: 'accounting', decimals: 2, symbol };
    case 'shortDate':
      return { kind: 'date', pattern: isJapanese ? 'yyyy/m/d' : 'm/d/yyyy' };
    case 'longDate':
      return { kind: 'date', pattern: isJapanese ? 'yyyy"年"m"月"d"日' : 'mmmm d, yyyy' };
    case 'time':
      return { kind: 'time', pattern: isJapanese ? 'H:MM' : 'h:MM AM/PM' };
    case 'percent':
      return { kind: 'percent', decimals: 0 };
    case 'fraction':
      return { kind: 'custom', pattern: '# ?/?' };
    case 'scientific':
      return { kind: 'scientific', decimals: 2 };
    case 'text':
      return { kind: 'text' };
    case 'more':
      return null;
  }
};
