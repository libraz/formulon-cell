import { currencySymbolForLocale, datePatternsForLocale } from '../format/locale.js';
import { formatNumber } from '../format/number-format.js';
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

const defaultCurrencySymbolForToolbar = currencySymbolForLocale;

export const numberFormatForAction = (
  action: NumberFormatAction,
  locale: string,
): NumFmt | null => {
  const patterns = datePatternsForLocale(locale);
  const symbol = defaultCurrencySymbolForToolbar(locale);
  switch (action) {
    case 'general':
      return { kind: 'general' };
    case 'fixed':
      return { kind: 'fixed', decimals: 2 };
    case 'currency':
      return { kind: 'currency', decimals: 2, symbol };
    case 'accounting':
      return { kind: 'accounting', decimals: 2, symbol };
    case 'shortDate':
      return { kind: 'date', pattern: patterns.short };
    case 'longDate':
      return { kind: 'date', pattern: patterns.long };
    case 'time':
      return { kind: 'time', pattern: patterns.time };
    case 'percent':
      return { kind: 'percent', decimals: 2 };
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

/** How `value` reads once the dropdown row `action` is applied; null for rows with no sample. */
export const numberFormatSample = (
  action: NumberFormatAction,
  value: number,
  locale: string,
): string | null => {
  if (action === 'general' || action === 'more') return null;
  const fmt = numberFormatForAction(action, locale);
  return fmt ? formatNumber(value, fmt, locale) : null;
};
