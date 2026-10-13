import type { ChartValueFormatter } from './types.js';

const finite = (value: number): number | null => (Number.isFinite(value) ? value : null);

const formatParts = (
  formatCode: string | undefined,
): {
  readonly valid: boolean;
  readonly percent: boolean;
  readonly grouping: boolean;
  readonly minimumFractionDigits: number;
  readonly maximumFractionDigits: number;
} => {
  const code = formatCode?.split(';', 1)[0]?.trim() ?? 'General';
  const valid = code.toLowerCase() === 'general' || /^[#0,]+(?:\.[#0]+)?%?$/.test(code);
  const percent = code.includes('%');
  const decimal = code.match(/\.([0#]+)/)?.[1];
  const minimumFractionDigits = decimal?.replace(/#/g, '').length ?? 0;
  const maximumFractionDigits = decimal?.length ?? 0;
  return {
    valid,
    percent,
    grouping: code.toLowerCase() === 'general' || code.includes(','),
    minimumFractionDigits,
    maximumFractionDigits: Math.max(minimumFractionDigits, maximumFractionDigits),
  };
};

/**
 * Small dependency-free fallback for the common Excel number formats used by
 * chart labels. Hosts may inject a formatter for the rest of the Excel format
 * language through layoutChart's formatValue callback.
 */
export const formatChartNumber = (
  value: number,
  formatCode: string | undefined,
  locale: string,
): string | null => {
  const finiteValue = finite(value);
  if (finiteValue === null) return null;
  const parts = formatParts(formatCode);
  if (!parts.valid) return formatChartNumber(finiteValue, undefined, locale);
  const code = formatCode?.split(';', 1)[0]?.trim().toLowerCase();
  const isGeneral = code === undefined || code === 'general';
  if (
    isGeneral &&
    (Math.abs(finiteValue) >= 1e12 || (Math.abs(finiteValue) > 0 && Math.abs(finiteValue) < 1e-9))
  ) {
    return finiteValue.toExponential(11).replace(/\.?(\d*?)0+e/, '$1e');
  }
  try {
    if (parts.percent) {
      return new Intl.NumberFormat(locale || 'en-US', {
        style: 'percent',
        useGrouping: parts.grouping,
        minimumFractionDigits: parts.minimumFractionDigits,
        maximumFractionDigits: parts.maximumFractionDigits,
      }).format(finiteValue);
    }
    const formatted = new Intl.NumberFormat(locale || 'en-US', {
      useGrouping: parts.grouping,
      minimumFractionDigits: parts.minimumFractionDigits,
      maximumFractionDigits: isGeneral ? 12 : parts.maximumFractionDigits,
    }).format(finiteValue);
    return isGeneral && formatted === '0' ? finiteValue.toExponential(11) : formatted;
  } catch {
    return String(finiteValue) + (parts.percent ? '%' : '');
  }
};

export const formatChartValue = (
  value: number,
  formatCode: string | undefined,
  locale: string,
  formatter?: ChartValueFormatter,
): string | null => {
  if (formatter) {
    const formatted = formatter(value, formatCode, locale);
    if (formatted !== null && formatted !== undefined) return formatted;
  }
  return formatChartNumber(value, formatCode, locale);
};
