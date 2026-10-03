import { splitFormulaArgs, splitFormulaArgsAllowEmpty } from '../splitter.js';
import type { FormulaCallParser, FormulaParserContext } from './context.js';

/** Exactly the parser entry points this family re-enters. */
export type ReferenceCallParserContext = Pick<
  FormulaParserContext,
  'parseFormulaOperand' | 'parseFormulaRangeArg'
>;

/** Reference and lookup calls: FORMULATEXT, CELL, SHEET(S), ROW/COLUMN,
 *  ROWS/COLUMNS/AREAS, MATCH/XMATCH, OFFSET, INDIRECT, INDEX and the LOOKUP family. */
export function createReferenceCallParser(ctx: ReferenceCallParserContext): FormulaCallParser {
  const { parseFormulaOperand, parseFormulaRangeArg } = ctx;
  return (fn, aggregate, sheetIndex) => {
    if (fn === 'FORMULATEXT') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const formulaRef = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (formulaRef) return { kind: 'formula-text', ref: formulaRef };
      }
    }
    if (fn === 'CELL') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1 || args?.length === 2) {
        const infoType = parseFormulaOperand(args[0] ?? '', sheetIndex);
        if (!infoType) return null;
        if (args.length === 1) return { kind: 'cell-info', infoType };
        const ref = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (ref) return { kind: 'cell-info', infoType, ref };
      }
    }
    if (fn === 'SHEET' || fn === 'SHEETS') {
      const rawArgs = aggregate[2] ?? '';
      if (rawArgs.trim() === '' && fn === 'SHEET') return { kind: 'sheet-info', fn };
      const args = splitFormulaArgs(rawArgs);
      if (args?.length === 1) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'sheet-info', fn, range };
      }
    }
    if (fn === 'ROW' || fn === 'COLUMN') {
      const rawArgs = aggregate[2] ?? '';
      if (rawArgs.trim() === '') return { kind: 'position', fn };
      const args = splitFormulaArgs(rawArgs);
      if (args?.length === 1) {
        const ref = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (ref) return { kind: 'position', fn, ref };
      }
    }
    if (fn === 'ROWS' || fn === 'COLUMNS' || fn === 'AREAS') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 1) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        if (range) return { kind: 'range-dimension', fn, range };
      }
    }
    if (fn === 'MATCH') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (lookup && range) {
          if (args.length === 2) return { kind: 'match', lookup, range };
          const matchType = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (matchType) return { kind: 'match', lookup, range, matchType };
        }
      }
    }
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
          const width =
            args.length >= 5 ? parseFormulaOperand(args[4] ?? '', sheetIndex) : undefined;
          if (args.length >= 5 && !width) return null;
          return {
            kind: 'offset',
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
        if (args.length === 1) return { kind: 'indirect', refText };
        const a1 = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (a1) return { kind: 'indirect', refText, a1 };
      }
    }
    if (fn === 'XMATCH') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 2 && args.length <= 4) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (lookup && range) {
          if (args.length === 2) return { kind: 'xmatch', lookup, range };
          const rawMatchMode = args[2] ?? '';
          const matchMode =
            rawMatchMode.trim() === '' ? undefined : parseFormulaOperand(rawMatchMode, sheetIndex);
          if (rawMatchMode.trim() !== '' && !matchMode) return null;
          if (args.length === 3) {
            return { kind: 'xmatch', lookup, range, ...(matchMode ? { matchMode } : {}) };
          }
          const searchMode = parseFormulaOperand(args[3] ?? '', sheetIndex);
          if (searchMode) {
            return {
              kind: 'xmatch',
              lookup,
              range,
              ...(matchMode ? { matchMode } : {}),
              searchMode,
            };
          }
        }
      }
    }
    if (fn === 'INDEX') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const range = parseFormulaRangeArg(args[0] ?? '', sheetIndex);
        const row = parseFormulaOperand(args[1] ?? '', sheetIndex);
        if (range && row) {
          if (args.length === 2) return { kind: 'index', range, row };
          const col = parseFormulaOperand(args[2] ?? '', sheetIndex);
          if (col) return { kind: 'index', range, row, col };
        }
      }
    }
    if (fn === 'VLOOKUP' || fn === 'HLOOKUP') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 3 || args?.length === 4) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const range = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const index = parseFormulaOperand(args[2] ?? '', sheetIndex);
        const rangeLookup =
          args.length === 3
            ? { kind: 'literal' as const, value: { kind: 'bool' as const, value: true } }
            : parseFormulaOperand(args[3] ?? '', sheetIndex);
        if (lookup && range && index && rangeLookup) {
          return { kind: 'lookup', fn, lookup, range, index, rangeLookup };
        }
      }
    }
    if (fn === 'XLOOKUP') {
      const args = splitFormulaArgsAllowEmpty(aggregate[2] ?? '');
      if (args && args.length >= 3 && args.length <= 6) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const lookupRange = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        const returnRange = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
        if (lookup && lookupRange && returnRange) {
          if (args.length === 3) return { kind: 'xlookup', lookup, lookupRange, returnRange };
          const rawIfNotFound = args[3] ?? '';
          const ifNotFound =
            rawIfNotFound.trim() === ''
              ? undefined
              : parseFormulaOperand(rawIfNotFound, sheetIndex);
          if (rawIfNotFound.trim() !== '' && !ifNotFound) return null;
          if (args.length === 4) {
            return {
              kind: 'xlookup',
              lookup,
              lookupRange,
              returnRange,
              ...(ifNotFound ? { ifNotFound } : {}),
            };
          }
          const rawMatchMode = args[4] ?? '';
          const matchMode =
            rawMatchMode.trim() === '' ? undefined : parseFormulaOperand(rawMatchMode, sheetIndex);
          if (rawMatchMode.trim() !== '' && !matchMode) return null;
          if (args.length === 5) {
            return {
              kind: 'xlookup',
              lookup,
              lookupRange,
              returnRange,
              ...(ifNotFound ? { ifNotFound } : {}),
              ...(matchMode ? { matchMode } : {}),
            };
          }
          const searchMode = parseFormulaOperand(args[5] ?? '', sheetIndex);
          if (searchMode) {
            return {
              kind: 'xlookup',
              lookup,
              lookupRange,
              returnRange,
              ...(ifNotFound ? { ifNotFound } : {}),
              ...(matchMode ? { matchMode } : {}),
              searchMode,
            };
          }
        }
      }
    }
    if (fn === 'LOOKUP') {
      const args = splitFormulaArgs(aggregate[2] ?? '');
      if (args?.length === 2 || args?.length === 3) {
        const lookup = parseFormulaOperand(args[0] ?? '', sheetIndex);
        const lookupRange = parseFormulaRangeArg(args[1] ?? '', sheetIndex);
        if (lookup && lookupRange) {
          if (args.length === 2) return { kind: 'vector-lookup', lookup, lookupRange };
          const resultRange = parseFormulaRangeArg(args[2] ?? '', sheetIndex);
          if (resultRange) return { kind: 'vector-lookup', lookup, lookupRange, resultRange };
        }
      }
    }
    return undefined;
  };
}
