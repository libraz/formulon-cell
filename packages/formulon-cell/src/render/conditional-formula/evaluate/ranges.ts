import { addrKey } from '../../../engine/address.js';
import type { CellValue } from '../../../engine/types.js';
import { aggregateResult, aggregateValueA } from '../aggregation.js';
import { readLogical, readNumber, textValue } from '../coercion.js';
import { MAX_FORMULA_AGGREGATE_CELLS, parseA1Range, parseR1C1Range } from '../parser.js';
import type {
  FormulaAggregateArg,
  FormulaAggregateName,
  FormulaRangeArg,
  FormulaRangeOperand,
  ParsedA1Range,
  ParsedRef,
} from '../types.js';
import type { FormulaReaderContext } from './context.js';

/**
 * Range resolution and cell collection shared by the function families:
 * relative refs resolve against the rule anchor, OFFSET/INDIRECT ranges are
 * read through `readOperand`, and every range is capped at
 * `MAX_FORMULA_AGGREGATE_CELLS`.
 */
export function createRangeReader(ctx: FormulaReaderContext) {
  const { state, sheet, anchorRow, anchorCol, readOperand } = ctx;

  const resolveRef = (ref: ParsedRef, rowOffset: number, colOffset: number): [number, number] => [
    ref.absRow ? ref.row : ref.row + rowOffset,
    ref.absCol ? ref.col : ref.col + colOffset,
  ];
  const rangeAggregateStats = (
    range: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): { values: number[]; valuesA: number[]; countA: number; countBlank: number } | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return null;
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    const valuesA: number[] = [];
    let countA = 0;
    let countBlank = 0;
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (value.kind === 'blank') countBlank += 1;
        else countA += 1;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
        const valueA = aggregateValueA(value);
        if (valueA !== null) valuesA.push(valueA);
      }
    }
    return { values, valuesA, countA, countBlank };
  };
  const aggregateRange = (
    range: FormulaRangeArg,
    fn: FormulaAggregateName,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const stats = rangeAggregateStats(range, rowOffset, colOffset);
    if (!stats) return { kind: 'error', code: 15, text: '#VALUE!' };
    return aggregateResult(fn, stats.values, stats.valuesA, stats.countA, stats.countBlank);
  };
  const aggregateArgs = (
    args: FormulaAggregateArg[],
    fn: FormulaAggregateName,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const values: number[] = [];
    const valuesA: number[] = [];
    let countA = 0;
    let countBlank = 0;
    for (const arg of args) {
      if (arg.kind === 'range' || arg.kind === 'dynamic-range') {
        const bounds = formulaRangeArgBounds(arg, rowOffset, colOffset);
        const stats = bounds ? rangeAggregateStatsFromBounds(bounds) : null;
        if (!stats) return { kind: 'error', code: 15, text: '#VALUE!' };
        values.push(...stats.values);
        valuesA.push(...stats.valuesA);
        countA += stats.countA;
        countBlank += stats.countBlank;
        continue;
      }
      const value = readOperand(arg.operand, rowOffset, colOffset);
      if (value.kind === 'blank') countBlank += 1;
      else countA += 1;
      if (value.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
      const valueA = aggregateValueA(value);
      if (valueA !== null) valuesA.push(valueA);
    }
    return aggregateResult(fn, values, valuesA, countA, countBlank);
  };
  const rangeBounds = (
    range: ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } => {
    const [startRow, startCol] = resolveRef(range.start, rowOffset, colOffset);
    const [endRow, endCol] = resolveRef(range.end, rowOffset, colOffset);
    const r0 = Math.min(startRow, endRow);
    const r1 = Math.max(startRow, endRow);
    const c0 = Math.min(startCol, endCol);
    const c1 = Math.max(startCol, endCol);
    const width = c1 - c0 + 1;
    const height = r1 - r0 + 1;
    return { r0, r1, c0, c1, width, height };
  };
  const fixedRangeBounds = (
    range: ParsedA1Range,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } => {
    const r0 = Math.min(range.start.row, range.end.row);
    const r1 = Math.max(range.start.row, range.end.row);
    const c0 = Math.min(range.start.col, range.end.col);
    const c1 = Math.max(range.start.col, range.end.col);
    return { r0, r1, c0, c1, width: c1 - c0 + 1, height: r1 - r0 + 1 };
  };
  const validRangeBounds = (bounds: { width: number; height: number }): boolean =>
    bounds.width > 0 &&
    bounds.height > 0 &&
    bounds.width * bounds.height <= MAX_FORMULA_AGGREGATE_CELLS;
  const dynamicRangeBounds = (
    range: FormulaRangeOperand,
    rowOffset: number,
    colOffset: number,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } | null => {
    if (range.kind === 'offset-range') {
      const rows = readNumber(readOperand(range.rows, rowOffset, colOffset));
      const cols = readNumber(readOperand(range.cols, rowOffset, colOffset));
      const rawHeight = range.height
        ? readNumber(readOperand(range.height, rowOffset, colOffset))
        : null;
      const rawWidth = range.width
        ? readNumber(readOperand(range.width, rowOffset, colOffset))
        : null;
      if (
        rows === null ||
        cols === null ||
        (range.height && rawHeight === null) ||
        (range.width && rawWidth === null)
      ) {
        return null;
      }
      const base = formulaRangeArgBounds(range.reference, rowOffset, colOffset);
      if (!base || !validRangeBounds(base)) return null;
      const height = rawHeight === null ? base.height : Math.trunc(rawHeight);
      const width = rawWidth === null ? base.width : Math.trunc(rawWidth);
      const r0 = base.r0 + Math.trunc(rows);
      const c0 = base.c0 + Math.trunc(cols);
      const bounds = { r0, r1: r0 + height - 1, c0, c1: c0 + width - 1, width, height };
      return bounds.r0 < 0 ||
        bounds.c0 < 0 ||
        bounds.r1 > 1048575 ||
        bounds.c1 > 16383 ||
        !validRangeBounds(bounds)
        ? null
        : bounds;
    }
    const refText = textValue(readOperand(range.refText, rowOffset, colOffset));
    const a1 = range.a1 ? readLogical(readOperand(range.a1, rowOffset, colOffset)) : true;
    if (refText === null || a1 === null) return null;
    const parsed = a1
      ? parseA1Range(refText, sheet)
      : parseR1C1Range(refText, sheet, anchorRow + rowOffset, anchorCol + colOffset);
    if (!parsed) return null;
    const bounds = fixedRangeBounds(parsed);
    return bounds.r0 < 0 ||
      bounds.c0 < 0 ||
      bounds.r1 > 1048575 ||
      bounds.c1 > 16383 ||
      !validRangeBounds(bounds)
      ? null
      : bounds;
  };
  const formulaRangeArgBounds = (
    arg: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { r0: number; r1: number; c0: number; c1: number; width: number; height: number } | null =>
    !('kind' in arg)
      ? rangeBounds(arg, rowOffset, colOffset)
      : arg.kind === 'range'
        ? rangeBounds(arg.range, rowOffset, colOffset)
        : dynamicRangeBounds(arg.range, rowOffset, colOffset);
  const singleCellRefPosition = (
    ref: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): { row: number; col: number } | null => {
    const bounds = formulaRangeArgBounds(ref, rowOffset, colOffset);
    if (bounds?.width !== 1 || bounds.height !== 1) return null;
    return { row: bounds.r0, col: bounds.c0 };
  };
  const rangeAggregateStatsFromBounds = (bounds: {
    r0: number;
    r1: number;
    c0: number;
    c1: number;
    width: number;
    height: number;
  }): { values: number[]; valuesA: number[]; countA: number; countBlank: number } | null => {
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    const valuesA: number[] = [];
    let countA = 0;
    let countBlank = 0;
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (value.kind === 'blank') countBlank += 1;
        else countA += 1;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
        const valueA = aggregateValueA(value);
        if (valueA !== null) valuesA.push(valueA);
      }
    }
    return { values, valuesA, countA, countBlank };
  };
  const numericValuesInBounds = (bounds: {
    r0: number;
    r1: number;
    c0: number;
    c1: number;
    width: number;
    height: number;
  }): number[] | null => {
    const { r0, r1, c0, c1 } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value;
        if (value?.kind === 'number' && Number.isFinite(value.value)) values.push(value.value);
      }
    }
    return values;
  };
  const numericValuesFromArgs = (
    args: FormulaAggregateArg[],
    rowOffset: number,
    colOffset: number,
  ): number[] | null => {
    const values: number[] = [];
    for (const arg of args) {
      if (arg.kind === 'range' || arg.kind === 'dynamic-range') {
        const bounds = formulaRangeArgBounds(arg, rowOffset, colOffset);
        const rangeValues = bounds ? numericValuesInBounds(bounds) : null;
        if (rangeValues === null) return null;
        values.push(...rangeValues);
        continue;
      }
      const value = readNumber(readOperand(arg.operand, rowOffset, colOffset));
      if (value === null) return null;
      values.push(value);
    }
    return values;
  };
  const numericValuesInFormulaRangeArg = (
    range: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): number[] | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    return bounds ? numericValuesInBounds(bounds) : null;
  };
  const numericValuesInRangeWithShape = (
    range: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { values: number[]; width: number; height: number } | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return null;
    const { r0, r1, c0, c1, width, height } = bounds;
    if (!validRangeBounds(bounds)) return null;
    const values: number[] = [];
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value;
        if (value?.kind !== 'number' || !Number.isFinite(value.value)) return null;
        values.push(value.value);
      }
    }
    return { values, width, height };
  };
  const numericPairsInRanges = (
    left: FormulaRangeArg | ParsedA1Range,
    right: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): { xs: number[]; ys: number[] } | null => {
    const leftBounds = formulaRangeArgBounds(left, rowOffset, colOffset);
    const rightBounds = formulaRangeArgBounds(right, rowOffset, colOffset);
    if (
      !leftBounds ||
      !rightBounds ||
      !validRangeBounds(leftBounds) ||
      !validRangeBounds(rightBounds) ||
      leftBounds.width !== rightBounds.width ||
      leftBounds.height !== rightBounds.height
    ) {
      return null;
    }
    const xs: number[] = [];
    const ys: number[] = [];
    for (let r = 0; r < leftBounds.height; r += 1) {
      for (let c = 0; c < leftBounds.width; c += 1) {
        const leftValue = state.data.cells.get(
          addrKey({ sheet, row: leftBounds.r0 + r, col: leftBounds.c0 + c }),
        )?.value;
        const rightValue = state.data.cells.get(
          addrKey({ sheet, row: rightBounds.r0 + r, col: rightBounds.c0 + c }),
        )?.value;
        if (
          leftValue?.kind === 'number' &&
          Number.isFinite(leftValue.value) &&
          rightValue?.kind === 'number' &&
          Number.isFinite(rightValue.value)
        ) {
          xs.push(leftValue.value);
          ys.push(rightValue.value);
        }
      }
    }
    return { xs, ys };
  };
  return {
    resolveRef,
    aggregateRange,
    aggregateArgs,
    validRangeBounds,
    formulaRangeArgBounds,
    singleCellRefPosition,
    numericValuesInBounds,
    numericValuesFromArgs,
    numericValuesInFormulaRangeArg,
    numericValuesInRangeWithShape,
    numericPairsInRanges,
  };
}

export type RangeReader = ReturnType<typeof createRangeReader>;
