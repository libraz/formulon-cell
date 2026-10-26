import type { ChartValueFormatter } from '../charts/types.js';
import { formatNumber } from '../format/number-format.js';

/** Keep spreadsheet format codes outside the independently reusable renderer.
 *  `date1904` is read per call so one formatter follows the live workbook. */
export const createSessionChartFormatter =
  (date1904: () => boolean): ChartValueFormatter =>
  (value, code, locale) =>
    formatNumber(
      value,
      code && code.trim().toLowerCase() !== 'general'
        ? { kind: 'custom', pattern: code }
        : undefined,
      locale,
      date1904(),
    );
