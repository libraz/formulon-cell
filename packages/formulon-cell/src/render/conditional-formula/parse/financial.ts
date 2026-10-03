import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaAggregateArg } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type FinancialCallParserContext = Pick<
  FormulaParserContext,
  'parseFormulaOperand' | 'parseFormulaRangeArg' | 'parseFormulaAggregateArg'
>;

/** Cash-flow calls: FVSCHEDULE, NPV, MIRR, XNPV, XIRR, IRR. Scalar financial
 *  functions parse through the numeric family. */
export function createFinancialCallParser(ctx: FinancialCallParserContext): FormulaCallParser {
  const { parseFormulaOperand, parseFormulaRangeArg, parseFormulaAggregateArg } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (fn === 'FVSCHEDULE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const principal = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const schedule = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (principal && schedule) return { kind: 'fv-schedule', principal, schedule };
      }
    }
    if (fn === 'NPV') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2) {
        const rate = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!rate) return null;
        const values: FormulaAggregateArg[] = [];
        for (const arg of args.slice(1)) {
          const valueArg = parseFormulaAggregateArg(arg, sheetIndex);
          if (!valueArg) return null;
          values.push(valueArg);
        }
        return { kind: 'npv', rate, values };
      }
    }
    if (fn === 'MIRR') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const financeRate = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const reinvestRate = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (values && financeRate && reinvestRate) {
          return { kind: 'mirr', values, financeRate, reinvestRate };
        }
      }
    }
    if (fn === 'XNPV') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const rate = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const values = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const dates = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
        if (rate && values && dates) return { kind: 'xnpv', rate, values, dates };
      }
    }
    if (fn === 'XIRR') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const dates = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (values && dates) {
          if (args.length === 2 || (args[2] ?? '').trim() === '') {
            return { kind: 'xirr', values, dates };
          }
          const guess = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (guess) return { kind: 'xirr', values, dates, guess };
        }
      }
    }
    if (fn === 'IRR') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const values = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (values) {
          if (args.length === 1 || (args[1] ?? '').trim() === '') {
            return { kind: 'irr', values };
          }
          const guess = parseFormulaOperand(args[1] ?? '', sheetIndex);
          if (guess) return { kind: 'irr', values, guess };
        }
      }
    }
    return undefined;
  };
}
