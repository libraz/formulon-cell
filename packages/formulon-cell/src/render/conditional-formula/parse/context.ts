import type {
  FormulaAggregateArg,
  FormulaCondition,
  FormulaOperand,
  FormulaRangeArg,
} from '../types.js';

/**
 * Recursive entry points a per-family call parser may re-enter for its
 * arguments. They live in `parser.ts`; passing them in keeps the family
 * modules free of a runtime import back into it. Each family narrows with `Pick`.
 */
export interface FormulaParserContext {
  readonly parseFormulaOperand: (raw: string, sheetIndex: number) => FormulaOperand | null;
  readonly parseFormulaRangeArg: (raw: string, sheetIndex: number) => FormulaRangeArg | null;
  readonly parseFormulaAggregateArg: (
    raw: string,
    sheetIndex: number,
  ) => FormulaAggregateArg | null;
  readonly parseFormulaCondition: (raw: string, sheetIndex: number) => FormulaCondition | null;
}

/**
 * Parses one `NAME(args)` call for the names its family owns. Returns
 * `undefined` when the name is not handled (or its arguments do not fit), so the
 * dispatcher moves on; `null` is a definitive parse failure.
 */
export type FormulaCallParser = (
  fn: string,
  aggregate: RegExpMatchArray,
  sheetIndex: number,
) => FormulaOperand | null | undefined;
