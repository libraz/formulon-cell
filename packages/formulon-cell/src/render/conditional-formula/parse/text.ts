import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaOperand } from '../types.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type TextCallParserContext = Pick<FormulaParserContext, 'parseFormulaOperand'>;

/** Text calls: LEN, HYPERLINK, search/slice/concat/substitute, TEXTBEFORE/AFTER,
 *  TEXTJOIN, case transforms, EXACT, TEXT/DOLLAR/FIXED, VALUE-style coercions, N/T. */
export function createTextCallParser(ctx: TextCallParserContext): FormulaCallParser {
  const { parseFormulaOperand } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (fn === 'LEN') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'text-length', value };
      }
    }
    if (fn === 'HYPERLINK') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const link = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const friendlyName =
          args.length === 2 ? parseFormulaOperand(args[1] ?? '', sheetIndex) : undefined;
        if (link && friendlyName !== null) return { kind: 'hyperlink', link, friendlyName };
      }
    }
    if (fn === 'SEARCH' || fn === 'FIND') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const needle = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const haystack = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (needle && haystack) {
          if (args.length === 2) return { kind: 'text-search', fn, needle, haystack };
          const start =
            (args[2] ?? '').trim() === ''
              ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } }
              : parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (start) return { kind: 'text-search', fn, needle, haystack, start };
        }
      }
    }
    if (fn === 'LEFT' || fn === 'RIGHT') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const count =
          (args[1] ?? '').trim() === ''
            ? { kind: 'literal' as const, value: { kind: 'number' as const, value: 1 } }
            : parseFormulaOperand(args[1] ?? '1', sheetIndex);
        if (value && count) return { kind: 'text-slice', fn, value, count };
      }
    }
    if (fn === 'MID') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const start = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const count = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (value && start && count) return { kind: 'text-slice', fn, value, start, count };
      }
    }
    if (fn === 'CONCATENATE' || fn === 'CONCAT') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length > 0) {
        const values = args.map((arg) =>
          arg.trim() === ''
            ? { kind: 'literal' as const, value: { kind: 'blank' as const } }
            : parseFormulaOperand(arg, sheetIndex),
        );
        if (values.every((value) => value !== null)) {
          return { kind: 'text-concat-function', values: values as FormulaOperand[] };
        }
      }
    }
    if (fn === 'SUBSTITUTE') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args?.length === 3 || args?.length === 4) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const oldText = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const newText = parseFormulaOperand(args[2] ?? '', sheetIndex);
        if (value && oldText && newText) {
          if (args.length === 3 || (args[3] ?? '').trim() === '') {
            return { kind: 'text-substitute', value, oldText, newText };
          }
          const instance = parseFormulaOperand(args[3] ?? '', sheetIndex);
          if (instance) return { kind: 'text-substitute', value, oldText, newText, instance };
        }
      }
    }
    if (fn === 'REPLACE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 4) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const start = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const count = parseFormulaOperand(args[2] ?? '', sheetIndex);
        const newText = parseFormulaOperand(args[3] ?? '', sheetIndex);
        if (value && start && count && newText) {
          return { kind: 'text-replace', value, start, count, newText };
        }
      }
    }
    if (fn === 'REPT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const count = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (value && count) return { kind: 'text-repeat', value, count };
      }
    }
    if (fn === 'TEXTBEFORE' || fn === 'TEXTAFTER') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (
        args &&
        args.length >= 2 &&
        args.length <= 6 &&
        (args[0] ?? '').trim() !== '' &&
        (args[1] ?? '').trim() !== ''
      ) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const delimiter = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const instance =
          args.length >= 3 && (args[2] ?? '').trim() !== ''
            ? parseFormulaOperand(args[2] ?? '', sheetIndex)
            : undefined;
        const matchMode =
          args.length >= 4 && (args[3] ?? '').trim() !== ''
            ? parseFormulaOperand(args[3] ?? '', sheetIndex)
            : undefined;
        const matchEnd =
          args.length >= 5 && (args[4] ?? '').trim() !== ''
            ? parseFormulaOperand(args[4] ?? '', sheetIndex)
            : undefined;
        const ifNotFound =
          args.length >= 6 && (args[5] ?? '').trim() !== ''
            ? parseFormulaOperand(args[5] ?? '', sheetIndex)
            : undefined;
        if (
          value &&
          delimiter &&
          instance !== null &&
          matchMode !== null &&
          matchEnd !== null &&
          ifNotFound !== null
        ) {
          return {
            kind: 'text-before-after',
            fn,
            value,
            delimiter,
            instance,
            matchMode,
            matchEnd,
            ifNotFound,
          };
        }
      }
    }
    if (fn === 'TEXTJOIN') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 3) {
        const delimiter = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const ignoreEmpty = parseFormulaOperand(args[1] ?? '', sheetIndex);
        const values = args
          .slice(2)
          .map((arg) =>
            arg.trim() === ''
              ? { kind: 'literal' as const, value: { kind: 'blank' as const } }
              : parseFormulaOperand(arg, sheetIndex),
          );
        if (delimiter && ignoreEmpty && values.every((value) => value !== null)) {
          return {
            kind: 'text-join',
            delimiter,
            ignoreEmpty,
            values: values as FormulaOperand[],
          };
        }
      }
    }
    if (
      fn === 'LOWER' ||
      fn === 'UPPER' ||
      fn === 'TRIM' ||
      fn === 'CLEAN' ||
      fn === 'PROPER' ||
      fn === 'ENCODEURL'
    ) {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'text-transform', fn, value };
      }
    }
    if (fn === 'EXACT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const left = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const right = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (left && right) return { kind: 'text-exact', left, right };
      }
    }
    if (fn === 'TEXT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const pattern = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (value && pattern) return { kind: 'text-format', value, pattern };
      }
    }
    if (fn === 'DOLLAR' || fn === 'FIXED') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (
        args &&
        args.length >= 1 &&
        args.length <= (fn === 'DOLLAR' ? 2 : 3) &&
        (args[0] ?? '').trim() !== ''
      ) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const decimals =
          args.length >= 2 && (args[1] ?? '').trim() !== ''
            ? parseFormulaOperand(args[1] ?? '', sheetIndex)
            : undefined;
        const noCommas =
          fn === 'FIXED' && args.length >= 3 && (args[2] ?? '').trim() !== ''
            ? parseFormulaOperand(args[2] ?? '', sheetIndex)
            : undefined;
        if (value && decimals !== null && noCommas !== null) {
          return { kind: 'text-fixed-format', fn, value, decimals, noCommas };
        }
      }
    }
    if (fn === 'VALUE') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'text-value', value };
      }
    }
    if (fn === 'VALUETOTEXT') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 1 && args.length <= 2 && (args[0] ?? '').trim() !== '') {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const format =
          args.length >= 2 && (args[1] ?? '').trim() !== ''
            ? parseFormulaOperand(args[1] ?? '', sheetIndex)
            : undefined;
        if (value && format !== null) return { kind: 'value-to-text', value, format };
      }
    }
    if (fn === 'NUMBERVALUE') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 1 && args.length <= 3 && (args[0] ?? '').trim() !== '') {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const decimalSeparator =
          args.length >= 2 && (args[1] ?? '').trim() !== ''
            ? parseFormulaOperand(args[1] ?? '', sheetIndex)
            : undefined;
        const groupSeparator =
          args.length >= 3 && (args[2] ?? '').trim() !== ''
            ? parseFormulaOperand(args[2] ?? '', sheetIndex)
            : undefined;
        if (value && decimalSeparator !== null && groupSeparator !== null) {
          return { kind: 'text-number-value', value, decimalSeparator, groupSeparator };
        }
      }
    }
    if (fn === 'N' || fn === 'T') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const value = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (value) return { kind: 'scalar-coerce', fn, value };
      }
    }
    return undefined;
  };
}
