import { splitFormulaArgs } from '../splitter.js';
import type { FormulaOperand, FormulaRangeArg } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type CriteriaCallParserContext = Pick<
  FormulaParserContext,
  'parseFormulaOperand' | 'parseFormulaRangeArg'
>;

/** Criteria calls: COUNTIF(S), SUMIF(S), AVERAGEIF(S), MINIFS/MAXIFS. */
export function createCriteriaCallParser(ctx: CriteriaCallParserContext): FormulaCallParser {
  const { parseFormulaOperand, parseFormulaRangeArg } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (fn === 'COUNTIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && criteria) return { kind: 'countif', range, criteria };
      }
    }
    if (fn === 'COUNTIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length % 2 === 0) {
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'countifs', pairs };
      }
    }
    if (fn === 'SUMIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const sumRange = parseFormulaRangeArg(args[2] ?? args[0] ?? '', sheetIndex);
        if (range && criteria && sumRange) return { kind: 'sumif', range, criteria, sumRange };
      }
    }
    if (fn === 'AVERAGEIF') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const criteria = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const averageRange = parseFormulaRangeArg(args[2] ?? args[0] ?? '', sheetIndex);
        if (range && criteria && averageRange) {
          return { kind: 'averageif', range, criteria, averageRange };
        }
      }
    }
    if (fn === 'SUMIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length % 2 === 1) {
        const sumRange = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (!sumRange) return null;
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'sumifs', sumRange, pairs };
      }
    }
    if (fn === 'AVERAGEIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length % 2 === 1) {
        const averageRange = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (!averageRange) return null;
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'averageifs', averageRange, pairs };
      }
    }
    if (fn === 'MINIFS' || fn === 'MAXIFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length % 2 === 1) {
        const valueRange = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (!valueRange) return null;
        const pairs: { range: FormulaRangeArg; criteria: FormulaOperand }[] = [];
        for (let i = 1; i < args.length; i += 2) {
          const range = parseFormulaRangeArg(args[i] ?? '', sheetIndex);
          const criteria = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!range || !criteria) return null;
          pairs.push({ range, criteria });
        }
        return { kind: 'minmaxifs', fn, valueRange, pairs };
      }
    }
    return undefined;
  };
}
