import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaDateArg } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type DatesCallParserContext = Pick<
  FormulaParserContext,
  'parseFormulaOperand' | 'parseFormulaRangeArg'
>;

/** Date and time calls. */
export function createDatesCallParser(ctx: DatesCallParserContext): FormulaCallParser {
  const { parseFormulaOperand, parseFormulaRangeArg } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (
      fn === 'DATE' ||
      fn === 'YEAR' ||
      fn === 'MONTH' ||
      fn === 'DAY' ||
      fn === 'WEEKDAY' ||
      fn === 'WEEKNUM' ||
      fn === 'ISOWEEKNUM' ||
      fn === 'TODAY' ||
      fn === 'NOW' ||
      fn === 'TIME' ||
      fn === 'EDATE' ||
      fn === 'EOMONTH' ||
      fn === 'DAYS' ||
      fn === 'DAYS360' ||
      fn === 'DATEDIF' ||
      fn === 'YEARFRAC' ||
      fn === 'DATEVALUE' ||
      fn === 'TIMEVALUE' ||
      fn === 'NETWORKDAYS' ||
      fn === 'NETWORKDAYS.INTL' ||
      fn === 'WORKDAY' ||
      fn === 'WORKDAY.INTL' ||
      fn === 'HOUR' ||
      fn === 'MINUTE' ||
      fn === 'SECOND'
    ) {
      const args =
        fn === 'WEEKDAY' ||
        fn === 'WEEKNUM' ||
        fn === 'DAYS360' ||
        fn === 'YEARFRAC' ||
        fn === 'NETWORKDAYS.INTL' ||
        fn === 'WORKDAY.INTL'
          ? splitFormulaArgsAllowEmpty(aggregate[2] ?? '')
          : splitFormulaArgs(aggregate[2] ?? '');
      const isIntlBusinessDayFn = fn === 'NETWORKDAYS.INTL' || fn === 'WORKDAY.INTL';
      const validLength =
        fn === 'TODAY' || fn === 'NOW'
          ? (aggregate[2] ?? '').trim() === ''
          : fn === 'DATE' || fn === 'TIME'
            ? args?.length === 3
            : fn === 'EDATE' ||
                fn === 'EOMONTH' ||
                fn === 'DAYS' ||
                fn === 'NETWORKDAYS' ||
                fn === 'WORKDAY' ||
                fn === 'NETWORKDAYS.INTL' ||
                fn === 'WORKDAY.INTL'
              ? args?.length === 2 ||
                ((fn === 'NETWORKDAYS' || fn === 'WORKDAY') && args?.length === 3) ||
                (isIntlBusinessDayFn && (args?.length === 3 || args?.length === 4))
              : fn === 'DATEDIF'
                ? args?.length === 3
                : fn === 'DAYS360'
                  ? args?.length === 2 || args?.length === 3
                  : fn === 'YEARFRAC'
                    ? args?.length === 2 || args?.length === 3
                    : fn === 'DATEVALUE' || fn === 'TIMEVALUE'
                      ? args?.length === 1
                      : fn === 'WEEKDAY' || fn === 'WEEKNUM'
                        ? args?.length === 1 || args?.length === 2
                        : args?.length === 1;
      if (fn === 'TODAY' || fn === 'NOW') {
        if (validLength) return { kind: 'date-function', fn, args: [] };
      } else if (args && validLength) {
        if (
          (fn === 'WEEKDAY' || fn === 'WEEKNUM' || fn === 'DAYS360' || fn === 'YEARFRAC') &&
          (args[0] ?? '').trim() === ''
        ) {
          return null;
        }
        const operands = args.map((arg, index) => {
          if ((fn === 'WEEKDAY' || fn === 'WEEKNUM') && index === 1 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } };
          }
          if (fn === 'DAYS360' && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'bool' as const, value: false } };
          }
          if (fn === 'YEARFRAC' && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } };
          }
          if (isIntlBusinessDayFn && index === 2 && arg.trim() === '') {
            return { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } };
          }
          if (
            ((fn === 'NETWORKDAYS' || fn === 'WORKDAY') && index === 2) ||
            (isIntlBusinessDayFn && index === 3)
          ) {
            const range = parseFormulaRangeArg(arg, sheetIndex);
            if (range) return range;
          }
          return parseFormulaOperand(arg, sheetIndex);
        });
        if (operands.every((operand) => operand !== null)) {
          return { kind: 'date-function', fn, args: operands as FormulaDateArg[] };
        }
      }
    }
    return undefined;
  };
}
