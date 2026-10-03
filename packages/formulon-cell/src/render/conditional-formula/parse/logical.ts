import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaCondition, FormulaOperand } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type LogicalCallParserContext = Pick<
  FormulaParserContext,
  'parseFormulaOperand' | 'parseFormulaCondition'
>;

/** Error-handling and selection calls: NA, IFERROR/IFNA, CHOOSE, SWITCH, IF, IFS. */
export function createLogicalCallParser(ctx: LogicalCallParserContext): FormulaCallParser {
  const { parseFormulaOperand, parseFormulaCondition } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (fn === 'NA') {
      const rawArgs = aggregate[2] ?? '';
      if (rawArgs.trim() === '') {
        return { kind: 'literal', value: { kind: 'error', code: 6, text: '#N/A' } };
      }
    }
    if (fn === 'IFERROR' || fn === 'IFNA') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const fallback = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (value && fallback) return { kind: 'error-fallback', fn, value, fallback };
      }
    }
    if (fn === 'CHOOSE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2) {
        const index = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const choices = args.slice(1).map((arg) => parseFormulaOperand(arg, sheetIndex));
        if (index && choices.every((choice) => choice !== null)) {
          return { kind: 'choose', index, choices: choices as FormulaOperand[] };
        }
      }
    }
    if (fn === 'SWITCH') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 3) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!value) return null;
        const hasDefault = args.length % 2 === 0;
        const caseArgs = hasDefault ? args.slice(1, -1) : args.slice(1);
        const cases: { match: FormulaOperand; result: FormulaOperand }[] = [];
        for (let i = 0; i < caseArgs.length; i += 2) {
          const match = parseFormulaOperand(caseArgs[i] ?? '', sheetIndex);
          const result = parseFormulaOperand(caseArgs[i + 1] ?? '', sheetIndex);
          if (!match || !result) return null;
          cases.push({ match, result });
        }
        if (hasDefault) {
          const defaultValue = parseFormulaOperand(args[args.length - 1] ?? '', sheetIndex);
          if (defaultValue) return { kind: 'switch', value, cases, defaultValue };
        } else {
          return { kind: 'switch', value, cases };
        }
      }
    }
    if (fn === 'IF') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && (args.length === 2 || args.length === 3)) {
        const condition = parseFormulaCondition(args[0] ?? '', sheetIndex);
        const whenTrue =
          args[1] === ''
            ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } }
            : parseFormulaOperand(args[1] ?? '', sheetIndex);
        const whenFalse =
          args.length === 2
            ? { kind: 'literal' as const, value: { kind: 'bool' as const, value: false } }
            : args[2] === ''
              ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 0 } }
              : parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (condition && whenTrue && whenFalse)
          return { kind: 'if', condition, whenTrue, whenFalse };
      }
    }
    if (fn === 'IFS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length % 2 === 0) {
        const branches: { condition: FormulaCondition; result: FormulaOperand }[] = [];
        for (let i = 0; i < args.length; i += 2) {
          const condition = parseFormulaCondition(args[i] ?? '', sheetIndex);
          const result = parseFormulaOperand(args[i + 1] ?? '', sheetIndex);
          if (!condition || !result) return null;
          branches.push({ condition, result });
        }
        return { kind: 'ifs', branches };
      }
    }
    return undefined;
  };
}
