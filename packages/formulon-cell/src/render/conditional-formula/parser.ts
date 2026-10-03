import type { FormulaCallParser, FormulaParserContext } from './parse/context.js';
import { createCriteriaCallParser } from './parse/criteria.js';
import { createDatesCallParser } from './parse/dates.js';
import { createFinancialCallParser } from './parse/financial.js';
import { createLogicalCallParser } from './parse/logical.js';
import { createNumericCallParser } from './parse/numeric.js';
import { createReferenceCallParser } from './parse/reference.js';
import { createStatisticsCallParser } from './parse/statistics.js';
import { createTextCallParser } from './parse/text.js';
import { parseA1Range, parseA1Ref, parseR1C1Range, parseR1C1Ref } from './references.js';
import {
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
  splitFormulaArithmetic,
  splitFormulaComparison,
  stripOuterParens,
} from './splitter.js';
import type {
  FormulaAggregateArg,
  FormulaCondition,
  FormulaOperand,
  FormulaRangeArg,
  FormulaRangeOperand,
} from './types.js';

const MAX_FORMULA_AGGREGATE_CELLS = 10000;
const FORMULA_NUMBER_LITERAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const FORMULA_VALUE_NUMBER_LITERAL =
  /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/;

const PARSER_CONTEXT: FormulaParserContext = {
  parseFormulaOperand,
  parseFormulaRangeArg,
  parseFormulaAggregateArg,
  parseFormulaCondition,
};

/** Per-family call parsers. Their name sets are disjoint, so order only affects speed. */
const FUNCTION_CALL_PARSERS: readonly FormulaCallParser[] = [
  createStatisticsCallParser(PARSER_CONTEXT),
  createFinancialCallParser(PARSER_CONTEXT),
  createCriteriaCallParser(PARSER_CONTEXT),
  createTextCallParser(PARSER_CONTEXT),
  createReferenceCallParser(PARSER_CONTEXT),
  createNumericCallParser(PARSER_CONTEXT),
  createDatesCallParser(PARSER_CONTEXT),
  createLogicalCallParser(PARSER_CONTEXT),
];

function parseFormulaRangeOperand(raw: string, sheetIndex: number): FormulaRangeOperand | null {
  const body = stripOuterParens(raw.trim());
  const aggregate = body.match(/^([A-Za-z][A-Za-z0-9.]*)\s*\((.*)\)$/);
  if (!aggregate) return null;
  const fn = (aggregate[1] ?? '').toUpperCase();
  if (fn === 'OFFSET') {
    const args = splitFormulaArgs(aggregate[2] ?? '');
    if (args && args.length >= 3 && args.length <= 5) {
      const reference = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
      const rows = parseFormulaOperand(args[1] ?? '', sheetIndex);
      const cols = parseFormulaOperand(args[2] ?? '', sheetIndex);
      if (reference && rows && cols) {
        const height =
          args.length >= 4 ? parseFormulaOperand(args[3] ?? '', sheetIndex) : undefined;
        if (args.length >= 4 && !height) return null;
        const width = args.length >= 5 ? parseFormulaOperand(args[4] ?? '', sheetIndex) : undefined;
        if (args.length >= 5 && !width) return null;
        return {
          kind: 'offset-range',
          reference,
          rows,
          cols,
          ...(height ? { height } : {}),
          ...(width ? { width } : {}),
        };
      }
    }
  }
  if (fn === 'INDIRECT') {
    const args = splitFormulaArgs(aggregate[2] ?? '');
    if (args?.length === 1 || args?.length === 2) {
      const refText = parseFormulaOperand(args[0] ?? '', sheetIndex);
      if (!refText) return null;
      if (args.length === 1) return { kind: 'indirect-range', refText };
      const a1 = parseFormulaOperand(args[1] ?? '', sheetIndex);
      if (a1) return { kind: 'indirect-range', refText, a1 };
    }
  }
  return null;
}

function parseFormulaAggregateArg(raw: string, sheetIndex: number): FormulaAggregateArg | null {
  const rangeArg = parseFormulaRangeArg(raw, sheetIndex);
  if (rangeArg) return rangeArg;
  const operand = parseFormulaOperand(raw, sheetIndex);
  return operand ? { kind: 'operand', operand } : null;
}

function parseFormulaRangeArg(raw: string, sheetIndex: number): FormulaRangeArg | null {
  const range = parseA1Range(raw, sheetIndex);
  if (range) return { kind: 'range', range };
  const dynamicRange = parseFormulaRangeOperand(raw, sheetIndex);
  if (dynamicRange) return { kind: 'dynamic-range', range: dynamicRange };
  return null;
}

function parseFormulaOperand(raw: string, sheetIndex: number): FormulaOperand | null {
  const body = stripOuterParens(raw.trim());
  const ref = parseA1Ref(body, sheetIndex);
  if (ref) return { kind: 'ref', ref };
  const aggregate = body.match(/^([A-Za-z][A-Za-z0-9.]*)\s*\((.*)\)$/);
  if (aggregate) {
    const fn = (aggregate[1] ?? '').toUpperCase();
    if ((fn === 'TRUE' || fn === 'FALSE') && (aggregate[2] ?? '').trim() === '') {
      return { kind: 'literal', value: { kind: 'bool', value: fn === 'TRUE' } };
    }
    if (fn === 'AND' || fn === 'OR' || fn === 'NOT' || fn === 'XOR') {
      const condition = parseFormulaCondition(body, sheetIndex);
      if (condition) return { kind: 'condition-value', condition };
    }
    for (const parseCall of FUNCTION_CALL_PARSERS) {
      const parsed = parseCall(fn, aggregate, sheetIndex);
      if (parsed !== undefined) return parsed;
    }
  }
  const arithmetic = splitFormulaArithmetic(body);
  if (arithmetic) {
    const left = parseFormulaOperand(arithmetic.left, sheetIndex);
    const right = parseFormulaOperand(arithmetic.right, sheetIndex);
    if (left && right) return { kind: 'binary', op: arithmetic.op, left, right };
  }
  if (
    (body.startsWith('"') && body.endsWith('"')) ||
    (body.startsWith("'") && body.endsWith("'"))
  ) {
    return { kind: 'literal', value: { kind: 'text', value: body.slice(1, -1) } };
  }
  if (FORMULA_NUMBER_LITERAL.test(body)) {
    return { kind: 'literal', value: { kind: 'number', value: Number(body) } };
  }
  if (/^true$/i.test(body)) return { kind: 'literal', value: { kind: 'bool', value: true } };
  if (/^false$/i.test(body)) return { kind: 'literal', value: { kind: 'bool', value: false } };
  return null;
}

function parseFormulaCondition(raw: string, sheetIndex: number): FormulaCondition | null {
  const body = stripOuterParens(raw.trim());
  if (/^true$/i.test(body)) return { kind: 'bool', value: true };
  if (/^false$/i.test(body)) return { kind: 'bool', value: false };
  const comparison = splitFormulaComparison(body);
  if (comparison) {
    const left = parseFormulaOperand(comparison.left, sheetIndex);
    const right = parseFormulaOperand(comparison.right, sheetIndex);
    if (left && right) return { kind: 'comparison', left, op: comparison.op, right };
  }
  const fnCall = body.match(/^([A-Za-z]+)\s*\((.*)\)$/);
  if (fnCall) {
    const fn = (fnCall[1] ?? '').toUpperCase();
    if ((fn === 'TRUE' || fn === 'FALSE') && (fnCall[2] ?? '').trim() === '') {
      return { kind: 'bool', value: fn === 'TRUE' };
    }
    if (fn === 'AND' || fn === 'OR' || fn === 'NOT' || fn === 'XOR') {
      const args = splitFormulaArgs(fnCall[2] ?? '');
      if (args === null || args.length === 0 || (fn === 'NOT' && args.length !== 1)) return null;
      const conditions = args.map((arg) => parseFormulaCondition(arg, sheetIndex));
      if (conditions.some((condition) => condition === null)) return null;
      return { kind: 'logical', fn, args: conditions as FormulaCondition[] };
    }
    if (
      fn === 'ISBLANK' ||
      fn === 'ISERROR' ||
      fn === 'ISERR' ||
      fn === 'ISNA' ||
      fn === 'ISNUMBER' ||
      fn === 'ISTEXT' ||
      fn === 'ISLOGICAL' ||
      fn === 'ISNONTEXT' ||
      fn === 'ISFORMULA' ||
      fn === 'ISREF'
    ) {
      const args = splitFormulaArgs(fnCall[2] ?? '');
      if (args?.length !== 1) return null;
      if (fn === 'ISREF' || fn === 'ISFORMULA') {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'is', fn, value: range };
        if (fn === 'ISREF') return null;
      }
      const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
      if (!value) return null;
      return { kind: 'is', fn, value };
    }
  }
  const operand = parseFormulaOperand(body, sheetIndex);
  return operand ? { kind: 'operand', value: operand } : null;
}

export {
  FORMULA_NUMBER_LITERAL,
  FORMULA_VALUE_NUMBER_LITERAL,
  MAX_FORMULA_AGGREGATE_CELLS,
  parseA1Range,
  parseA1Ref,
  parseFormulaOperand,
  parseFormulaRangeArg,
  parseR1C1Range,
  parseR1C1Ref,
  splitFormulaArgs,
  splitFormulaArgsAllowEmpty,
  splitFormulaComparison,
  stripOuterParens,
};
