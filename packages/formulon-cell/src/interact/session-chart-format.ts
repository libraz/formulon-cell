import type { ChartValueFormatter } from '../charts/types.js';
import { formatNumber } from '../format/number-format.js';

/** Keep spreadsheet format codes outside the independently reusable renderer. */
export const formatSessionChartValue: ChartValueFormatter = (value, code, locale) =>
  formatNumber(
    value,
    code && code.trim().toLowerCase() !== 'general' ? { kind: 'custom', pattern: code } : undefined,
    locale,
  );
