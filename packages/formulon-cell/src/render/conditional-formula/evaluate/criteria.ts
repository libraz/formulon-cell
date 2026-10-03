import { addrKey } from '../../../engine/address.js';
import type { CellValue } from '../../../engine/types.js';
import { matchesCountIfCriteria } from '../matching.js';
import type { FormulaRangeArg } from '../types.js';
import type { FormulaReaderContext } from './context.js';
import type { RangeReader } from './ranges.js';

/** Exactly the reader members this family touches. */
export type CriteriaEvaluatorContext = Pick<FormulaReaderContext, 'state' | 'sheet'> &
  Pick<RangeReader, 'formulaRangeArgBounds' | 'validRangeBounds'>;

/** Criteria-filtered ranges: COUNTIF(S), SUMIF(S), AVERAGEIF(S), MINIFS/MAXIFS. */
export function createCriteriaEvaluator(ctx: CriteriaEvaluatorContext) {
  const { state, sheet, formulaRangeArgBounds, validRangeBounds } = ctx;
  const countMatchingRange = (
    range: FormulaRangeArg,
    criteria: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let count = 0;
    for (let r = bounds.r0; r <= bounds.r1; r += 1) {
      for (let c = bounds.c0; c <= bounds.c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (matchesCountIfCriteria(value, criteria)) count += 1;
      }
    }
    return { kind: 'number', value: count };
  };
  const countMatchingRanges = (
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (!nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    const first = bounds[0];
    if (!first) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (bounds.some((bound) => bound.width !== first.width || bound.height !== first.height)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let count = 0;
    for (let dr = 0; dr < first.height; dr += 1) {
      for (let dc = 0; dc < first.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof first;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (matches) count += 1;
      }
    }
    return { kind: 'number', value: count };
  };
  const sumMatchingRange = (
    range: FormulaRangeArg,
    criteria: CellValue,
    sumRange: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const criteriaBounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    const sumBounds = formulaRangeArgBounds(sumRange, rowOffset, colOffset);
    if (
      !criteriaBounds ||
      !sumBounds ||
      !validRangeBounds(criteriaBounds) ||
      !validRangeBounds(sumBounds)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (sumBounds.width !== criteriaBounds.width || sumBounds.height !== criteriaBounds.height) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    for (let dr = 0; dr < criteriaBounds.height; dr += 1) {
      for (let dc = 0; dc < criteriaBounds.width; dc += 1) {
        const criteriaValue = state.data.cells.get(
          addrKey({ sheet, row: criteriaBounds.r0 + dr, col: criteriaBounds.c0 + dc }),
        )?.value ?? { kind: 'blank' as const };
        if (!matchesCountIfCriteria(criteriaValue, criteria)) continue;
        const sumValue = state.data.cells.get(
          addrKey({ sheet, row: sumBounds.r0 + dr, col: sumBounds.c0 + dc }),
        )?.value;
        if (sumValue?.kind === 'number' && Number.isFinite(sumValue.value)) sum += sumValue.value;
      }
    }
    return { kind: 'number', value: sum };
  };
  const averageMatchingRange = (
    range: FormulaRangeArg,
    criteria: CellValue,
    averageRange: FormulaRangeArg,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const criteriaBounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    const averageBounds = formulaRangeArgBounds(averageRange, rowOffset, colOffset);
    if (
      !criteriaBounds ||
      !averageBounds ||
      !validRangeBounds(criteriaBounds) ||
      !validRangeBounds(averageBounds)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (
      averageBounds.width !== criteriaBounds.width ||
      averageBounds.height !== criteriaBounds.height
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    let count = 0;
    for (let dr = 0; dr < criteriaBounds.height; dr += 1) {
      for (let dc = 0; dc < criteriaBounds.width; dc += 1) {
        const criteriaValue = state.data.cells.get(
          addrKey({ sheet, row: criteriaBounds.r0 + dr, col: criteriaBounds.c0 + dc }),
        )?.value ?? { kind: 'blank' as const };
        if (!matchesCountIfCriteria(criteriaValue, criteria)) continue;
        const averageValue = state.data.cells.get(
          addrKey({ sheet, row: averageBounds.r0 + dr, col: averageBounds.c0 + dc }),
        )?.value;
        if (averageValue?.kind === 'number' && Number.isFinite(averageValue.value)) {
          sum += averageValue.value;
          count += 1;
        }
      }
    }
    return count > 0
      ? { kind: 'number', value: sum / count }
      : { kind: 'error', code: 1, text: '#DIV/0!' };
  };
  const sumMatchingRanges = (
    sumRange: FormulaRangeArg,
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const sumBounds = formulaRangeArgBounds(sumRange, rowOffset, colOffset);
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (
      !sumBounds ||
      !validRangeBounds(sumBounds) ||
      !nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    if (
      bounds.some((bound) => bound.width !== sumBounds.width || bound.height !== sumBounds.height)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    for (let dr = 0; dr < sumBounds.height; dr += 1) {
      for (let dc = 0; dc < sumBounds.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof sumBounds;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (!matches) continue;
        const sumValue = state.data.cells.get(
          addrKey({ sheet, row: sumBounds.r0 + dr, col: sumBounds.c0 + dc }),
        )?.value;
        if (sumValue?.kind === 'number' && Number.isFinite(sumValue.value)) sum += sumValue.value;
      }
    }
    return { kind: 'number', value: sum };
  };
  const averageMatchingRanges = (
    averageRange: FormulaRangeArg,
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const averageBounds = formulaRangeArgBounds(averageRange, rowOffset, colOffset);
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (
      !averageBounds ||
      !validRangeBounds(averageBounds) ||
      !nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    if (
      bounds.some(
        (bound) => bound.width !== averageBounds.width || bound.height !== averageBounds.height,
      )
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let sum = 0;
    let count = 0;
    for (let dr = 0; dr < averageBounds.height; dr += 1) {
      for (let dc = 0; dc < averageBounds.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof averageBounds;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (!matches) continue;
        const averageValue = state.data.cells.get(
          addrKey({ sheet, row: averageBounds.r0 + dr, col: averageBounds.c0 + dc }),
        )?.value;
        if (averageValue?.kind === 'number' && Number.isFinite(averageValue.value)) {
          sum += averageValue.value;
          count += 1;
        }
      }
    }
    return count > 0
      ? { kind: 'number', value: sum / count }
      : { kind: 'error', code: 1, text: '#DIV/0!' };
  };
  const minMaxMatchingRanges = (
    valueRange: FormulaRangeArg,
    fn: 'MINIFS' | 'MAXIFS',
    pairs: { range: FormulaRangeArg; criteria: CellValue }[],
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const valueBounds = formulaRangeArgBounds(valueRange, rowOffset, colOffset);
    const nullableBounds = pairs.map((pair) =>
      formulaRangeArgBounds(pair.range, rowOffset, colOffset),
    );
    if (
      !valueBounds ||
      !validRangeBounds(valueBounds) ||
      !nullableBounds.every((bound) => bound !== null && validRangeBounds(bound))
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = nullableBounds as NonNullable<(typeof nullableBounds)[number]>[];
    if (
      bounds.some(
        (bound) => bound.width !== valueBounds.width || bound.height !== valueBounds.height,
      )
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let result: number | null = null;
    for (let dr = 0; dr < valueBounds.height; dr += 1) {
      for (let dc = 0; dc < valueBounds.width; dc += 1) {
        const matches = pairs.every((pair, index) => {
          const bound = bounds[index] as typeof valueBounds;
          const value = state.data.cells.get(
            addrKey({ sheet, row: bound.r0 + dr, col: bound.c0 + dc }),
          )?.value ?? { kind: 'blank' as const };
          return matchesCountIfCriteria(value, pair.criteria);
        });
        if (!matches) continue;
        const value = state.data.cells.get(
          addrKey({ sheet, row: valueBounds.r0 + dr, col: valueBounds.c0 + dc }),
        )?.value;
        if (value?.kind !== 'number' || !Number.isFinite(value.value)) continue;
        result =
          result === null
            ? value.value
            : fn === 'MINIFS'
              ? Math.min(result, value.value)
              : Math.max(result, value.value);
      }
    }
    return result === null
      ? { kind: 'error', code: 1, text: '#DIV/0!' }
      : { kind: 'number', value: result };
  };
  return {
    countMatchingRange,
    countMatchingRanges,
    sumMatchingRange,
    averageMatchingRange,
    sumMatchingRanges,
    averageMatchingRanges,
    minMaxMatchingRanges,
  };
}
