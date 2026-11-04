import {
  currencySymbolForLocale,
  datePatternsForLocale,
  normalizeFormatLocale,
} from '../format/locale.js';
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

/** The long date as Excel's dropdown applies it: with the weekday where the locale pattern has none. */
const withWeekday = (long: string, locale: string): string => {
  if (long.includes('dddd')) return long;
  const language = normalizeFormatLocale(locale).split('-')[0];
  if (language === 'ja') return `${long} aaaa`;
  if (language === 'en') return `dddd, ${long}`;
  return long;
};

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
      return { kind: 'date', pattern: withWeekday(patterns.long, locale) };
    case 'time':
      return { kind: 'time', pattern: patterns.time.replace('MM', 'MM:ss') };
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
