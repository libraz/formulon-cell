import { addrKey, formatA1Cell, MAX_COL, MAX_ROW } from '../../../engine/address.js';
import type { CellValue } from '../../../engine/types.js';
import { readLogical, readNumber, textValue } from '../coercion.js';
import {
  approximateMatchIndex,
  approximateXmatchIndex,
  exactMatchValues,
  isApproximateLookupMode,
  isExactLookupMode,
} from '../matching.js';
import { parseA1Ref, parseR1C1Ref } from '../parser.js';
import type { FormulaRangeArg, ParsedA1Range } from '../types.js';
import type { FormulaReaderContext } from './context.js';
import type { RangeReader } from './ranges.js';

/** Exactly the reader members this family touches. */
export type LookupEvaluatorContext = Pick<
  FormulaReaderContext,
  'state' | 'sheet' | 'anchorRow' | 'anchorCol'
> &
  Pick<RangeReader, 'formulaRangeArgBounds' | 'validRangeBounds' | 'singleCellRefPosition'>;

/** Lookup and reference functions: MATCH/XMATCH, INDEX, OFFSET, INDIRECT,
 *  VLOOKUP/HLOOKUP/LOOKUP/XLOOKUP, CELL and SHEET/SHEETS. */
export function createLookupEvaluator(ctx: LookupEvaluatorContext) {
  const {
    state,
    sheet,
    anchorRow,
    anchorCol,
    formulaRangeArgBounds,
    validRangeBounds,
    singleCellRefPosition,
  } = ctx;
  const oneDimensionalValues = (
    range: FormulaRangeArg | ParsedA1Range,
    rowOffset: number,
    colOffset: number,
  ): CellValue[] | null => {
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return null;
    if (!validRangeBounds(bounds) || (bounds.width !== 1 && bounds.height !== 1)) return null;
    const vertical = bounds.width === 1;
    const count = vertical ? bounds.height : bounds.width;
    const values: CellValue[] = [];
    for (let i = 0; i < count; i += 1) {
      const row = vertical ? bounds.r0 + i : bounds.r0;
      const col = vertical ? bounds.c0 : bounds.c0 + i;
      values.push(state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? { kind: 'blank' });
    }
    return values;
  };
  const matchExactRange = (
    lookup: CellValue,
    range: FormulaRangeArg,
    matchType: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const matchTypeValue = matchType === null ? 1 : readNumber(matchType);
    if (matchTypeValue === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const matchTypeInt = Math.trunc(matchTypeValue);
    if (matchTypeInt === 1 || matchTypeInt === -1) {
      const values = oneDimensionalValues(range, rowOffset, colOffset);
      if (!values) return { kind: 'error', code: 15, text: '#VALUE!' };
      const index = approximateMatchIndex(lookup, values, matchTypeInt);
      return index === null
        ? { kind: 'error', code: 6, text: '#N/A' }
        : { kind: 'number', value: index + 1 };
    }
    if (matchType !== null) {
      const value = readNumber(matchType);
      if (value === null || value !== 0) return { kind: 'error', code: 6, text: '#N/A' };
    }
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds) || (bounds.width !== 1 && bounds.height !== 1)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let index = 1;
    for (let r = bounds.r0; r <= bounds.r1; r += 1) {
      for (let c = bounds.c0; c <= bounds.c1; c += 1) {
        const value = state.data.cells.get(addrKey({ sheet, row: r, col: c }))?.value ?? {
          kind: 'blank' as const,
        };
        if (exactMatchValues(lookup, value)) return { kind: 'number', value: index };
        index += 1;
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  const xmatchRange = (
    lookup: CellValue,
    range: FormulaRangeArg,
    matchMode: CellValue | null,
    searchMode: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const matchModeValue = matchMode === null ? 0 : readNumber(matchMode);
    const searchModeValue = searchMode === null ? 1 : readNumber(searchMode);
    if (matchModeValue === null || searchModeValue === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const matchModeInt = Math.trunc(matchModeValue);
    const searchModeInt = Math.trunc(searchModeValue);
    if (
      (matchModeInt !== 0 && matchModeInt !== 2 && matchModeInt !== -1 && matchModeInt !== 1) ||
      (searchModeInt !== 1 && searchModeInt !== -1)
    ) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    const values = oneDimensionalValues(range, rowOffset, colOffset);
    if (!values) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    if (matchModeInt === -1 || matchModeInt === 1) {
      const index = approximateXmatchIndex(lookup, values, matchModeInt);
      return index === null
        ? { kind: 'error', code: 6, text: '#N/A' }
        : { kind: 'number', value: index + 1 };
    }
    for (
      let i = searchModeInt === -1 ? values.length - 1 : 0;
      i >= 0 && i < values.length;
      i += searchModeInt
    ) {
      const candidate = values[i] as CellValue;
      if (exactMatchValues(lookup, candidate, matchModeInt === 2)) {
        return { kind: 'number', value: i + 1 };
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  const indexRange = (
    range: FormulaRangeArg,
    rowValue: CellValue,
    colValue: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rowNumber = readNumber(rowValue);
    const colNumber = colValue === null ? null : readNumber(colValue);
    if (rowNumber === null || (colValue !== null && colNumber === null)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const rowIndex = Math.trunc(rowNumber);
    const colIndex = colNumber === null ? null : Math.trunc(colNumber);
    if (rowIndex < 1 || (colIndex !== null && colIndex < 1)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    let targetRow: number;
    let targetCol: number;
    if (colIndex === null) {
      if (bounds.width === 1) {
        targetRow = bounds.r0 + rowIndex - 1;
        targetCol = bounds.c0;
      } else if (bounds.height === 1) {
        targetRow = bounds.r0;
        targetCol = bounds.c0 + rowIndex - 1;
      } else {
        return { kind: 'error', code: 15, text: '#VALUE!' };
      }
    } else {
      targetRow = bounds.r0 + rowIndex - 1;
      targetCol = bounds.c0 + colIndex - 1;
    }
    if (targetRow > bounds.r1 || targetCol > bounds.c1) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return (
      state.data.cells.get(addrKey({ sheet, row: targetRow, col: targetCol }))?.value ?? {
        kind: 'blank',
      }
    );
  };
  const offsetValue = (
    reference: FormulaRangeArg,
    rowsValue: CellValue,
    colsValue: CellValue,
    heightValue: CellValue | null,
    widthValue: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const rows = readNumber(rowsValue);
    const cols = readNumber(colsValue);
    const rawHeight = heightValue === null ? null : readNumber(heightValue);
    const rawWidth = widthValue === null ? null : readNumber(widthValue);
    if (
      rows === null ||
      cols === null ||
      (heightValue !== null && rawHeight === null) ||
      (widthValue !== null && rawWidth === null)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const bounds = formulaRangeArgBounds(reference, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const height = rawHeight === null ? bounds.height : Math.trunc(rawHeight);
    const width = rawWidth === null ? bounds.width : Math.trunc(rawWidth);
    if (height !== 1 || width !== 1) return { kind: 'error', code: 15, text: '#VALUE!' };
    const targetRow = bounds.r0 + Math.trunc(rows);
    const targetCol = bounds.c0 + Math.trunc(cols);
    if (targetRow < 0 || targetRow > MAX_ROW || targetCol < 0 || targetCol > MAX_COL) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return (
      state.data.cells.get(addrKey({ sheet, row: targetRow, col: targetCol }))?.value ?? {
        kind: 'blank',
      }
    );
  };
  const indirectValue = (
    refTextValue: CellValue,
    a1Value: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const refText = textValue(refTextValue);
    const a1 = a1Value === null ? true : readLogical(a1Value);
    if (refText === null || a1 === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const ref = a1
      ? parseA1Ref(refText, sheet)
      : parseR1C1Ref(refText, sheet, anchorRow + rowOffset, anchorCol + colOffset);
    if (!ref) return { kind: 'error', code: 15, text: '#VALUE!' };
    const row = ref.row;
    const col = ref.col;
    if (row < 0 || row > MAX_ROW || col < 0 || col > MAX_COL) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    return (
      state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? {
        kind: 'blank',
      }
    );
  };
  const tableLookup = (
    fn: 'VLOOKUP' | 'HLOOKUP',
    lookup: CellValue,
    range: FormulaRangeArg,
    indexValue: CellValue,
    rangeLookup: CellValue,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (!isExactLookupMode(rangeLookup) && !isApproximateLookupMode(rangeLookup)) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    const indexNumber = readNumber(indexValue);
    if (indexNumber === null) return { kind: 'error', code: 15, text: '#VALUE!' };
    const index = Math.trunc(indexNumber);
    if (index < 1) return { kind: 'error', code: 15, text: '#VALUE!' };
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds || !validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    const approximate = isApproximateLookupMode(rangeLookup);
    if (fn === 'VLOOKUP') {
      if (index > bounds.width) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (approximate) {
        const values: CellValue[] = [];
        for (let r = bounds.r0; r <= bounds.r1; r += 1) {
          values.push(
            state.data.cells.get(addrKey({ sheet, row: r, col: bounds.c0 }))?.value ?? {
              kind: 'blank',
            },
          );
        }
        const matchIndex = approximateMatchIndex(lookup, values, 1);
        if (matchIndex === null) return { kind: 'error', code: 6, text: '#N/A' };
        return (
          state.data.cells.get(
            addrKey({ sheet, row: bounds.r0 + matchIndex, col: bounds.c0 + index - 1 }),
          )?.value ?? { kind: 'blank' }
        );
      }
      for (let r = bounds.r0; r <= bounds.r1; r += 1) {
        const candidate = state.data.cells.get(addrKey({ sheet, row: r, col: bounds.c0 }))
          ?.value ?? {
          kind: 'blank' as const,
        };
        if (exactMatchValues(lookup, candidate)) {
          return (
            state.data.cells.get(addrKey({ sheet, row: r, col: bounds.c0 + index - 1 }))?.value ?? {
              kind: 'blank',
            }
          );
        }
      }
    } else {
      if (index > bounds.height) return { kind: 'error', code: 15, text: '#VALUE!' };
      if (approximate) {
        const values: CellValue[] = [];
        for (let c = bounds.c0; c <= bounds.c1; c += 1) {
          values.push(
            state.data.cells.get(addrKey({ sheet, row: bounds.r0, col: c }))?.value ?? {
              kind: 'blank',
            },
          );
        }
        const matchIndex = approximateMatchIndex(lookup, values, 1);
        if (matchIndex === null) return { kind: 'error', code: 6, text: '#N/A' };
        return (
          state.data.cells.get(
            addrKey({ sheet, row: bounds.r0 + index - 1, col: bounds.c0 + matchIndex }),
          )?.value ?? { kind: 'blank' }
        );
      }
      for (let c = bounds.c0; c <= bounds.c1; c += 1) {
        const candidate = state.data.cells.get(addrKey({ sheet, row: bounds.r0, col: c }))
          ?.value ?? {
          kind: 'blank' as const,
        };
        if (exactMatchValues(lookup, candidate)) {
          return (
            state.data.cells.get(addrKey({ sheet, row: bounds.r0 + index - 1, col: c }))?.value ?? {
              kind: 'blank',
            }
          );
        }
      }
    }
    return { kind: 'error', code: 6, text: '#N/A' };
  };
  const xlookupRange = (
    lookup: CellValue,
    lookupRange: FormulaRangeArg,
    returnRange: FormulaRangeArg,
    ifNotFound: CellValue | null,
    matchMode: CellValue | null,
    searchMode: CellValue | null,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const matchModeValue = matchMode === null ? 0 : readNumber(matchMode);
    const searchModeValue = searchMode === null ? 1 : readNumber(searchMode);
    if (matchModeValue === null || searchModeValue === null) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const matchModeInt = Math.trunc(matchModeValue);
    const searchModeInt = Math.trunc(searchModeValue);
    if (
      (matchModeInt !== 0 && matchModeInt !== 2 && matchModeInt !== -1 && matchModeInt !== 1) ||
      (searchModeInt !== 1 && searchModeInt !== -1)
    ) {
      return { kind: 'error', code: 6, text: '#N/A' };
    }
    const lookupBounds = formulaRangeArgBounds(lookupRange, rowOffset, colOffset);
    const returnBounds = formulaRangeArgBounds(returnRange, rowOffset, colOffset);
    if (
      !lookupBounds ||
      !returnBounds ||
      !validRangeBounds(lookupBounds) ||
      !validRangeBounds(returnBounds) ||
      (lookupBounds.width !== 1 && lookupBounds.height !== 1)
    ) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const vertical = lookupBounds.width === 1;
    const count = vertical ? lookupBounds.height : lookupBounds.width;
    if ((vertical && returnBounds.height < count) || (!vertical && returnBounds.width < count)) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const returnAt = (index: number): CellValue =>
      state.data.cells.get(
        addrKey({
          sheet,
          row: vertical ? returnBounds.r0 + index : returnBounds.r0,
          col: vertical ? returnBounds.c0 : returnBounds.c0 + index,
        }),
      )?.value ?? { kind: 'blank' };
    if (matchModeInt === -1 || matchModeInt === 1) {
      const values: CellValue[] = [];
      for (let i = 0; i < count; i += 1) {
        const row = vertical ? lookupBounds.r0 + i : lookupBounds.r0;
        const col = vertical ? lookupBounds.c0 : lookupBounds.c0 + i;
        values.push(state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? { kind: 'blank' });
      }
      const matchIndex = approximateXmatchIndex(lookup, values, matchModeInt);
      return matchIndex === null
        ? (ifNotFound ?? { kind: 'error', code: 6, text: '#N/A' })
        : returnAt(matchIndex);
    }
    for (let i = searchModeInt === -1 ? count - 1 : 0; i >= 0 && i < count; i += searchModeInt) {
      const lookupRow = vertical ? lookupBounds.r0 + i : lookupBounds.r0;
      const lookupCol = vertical ? lookupBounds.c0 : lookupBounds.c0 + i;
      const candidate = state.data.cells.get(addrKey({ sheet, row: lookupRow, col: lookupCol }))
        ?.value ?? {
        kind: 'blank' as const,
      };
      if (exactMatchValues(lookup, candidate, matchModeInt === 2)) {
        return returnAt(i);
      }
    }
    return ifNotFound ?? { kind: 'error', code: 6, text: '#N/A' };
  };
  const vectorLookup = (
    lookup: CellValue,
    lookupRange: FormulaRangeArg,
    resultRange: FormulaRangeArg | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const lookupValues = oneDimensionalValues(lookupRange, rowOffset, colOffset);
    if (!lookupValues) return { kind: 'error', code: 15, text: '#VALUE!' };
    const resultValues = resultRange
      ? oneDimensionalValues(resultRange, rowOffset, colOffset)
      : lookupValues;
    if (!resultValues || resultValues.length !== lookupValues.length) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const matchIndex = approximateMatchIndex(lookup, lookupValues, 1);
    return matchIndex === null
      ? { kind: 'error', code: 6, text: '#N/A' }
      : (resultValues[matchIndex] as CellValue);
  };
  const cellInfo = (
    infoType: CellValue,
    ref: FormulaRangeArg | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    const type = textValue(infoType)?.trim().toLowerCase();
    if (!type) return { kind: 'error', code: 15, text: '#VALUE!' };
    const position = ref ? singleCellRefPosition(ref, rowOffset, colOffset) : null;
    if (ref && !position) return { kind: 'error', code: 15, text: '#VALUE!' };
    const [row, col] = position
      ? [position.row, position.col]
      : [anchorRow + rowOffset, anchorCol + colOffset];
    if (row < 0 || row > MAX_ROW || col < 0 || col > MAX_COL) {
      return { kind: 'error', code: 15, text: '#VALUE!' };
    }
    const value = state.data.cells.get(addrKey({ sheet, row, col }))?.value ?? {
      kind: 'blank' as const,
    };
    if (type === 'address') {
      return { kind: 'text', value: formatA1Cell(row, col, true) };
    }
    if (type === 'row') return { kind: 'number', value: row + 1 };
    if (type === 'col') return { kind: 'number', value: col + 1 };
    if (type === 'contents') return value;
    if (type === 'type') {
      return {
        kind: 'text',
        value: value.kind === 'blank' ? 'b' : value.kind === 'text' ? 'l' : 'v',
      };
    }
    return { kind: 'error', code: 15, text: '#VALUE!' };
  };
  const sheetInfo = (
    fn: 'SHEET' | 'SHEETS',
    range: FormulaRangeArg | undefined,
    rowOffset: number,
    colOffset: number,
  ): CellValue => {
    if (!range) return { kind: 'number', value: sheet + 1 };
    const bounds = formulaRangeArgBounds(range, rowOffset, colOffset);
    if (!bounds) return { kind: 'error', code: 15, text: '#VALUE!' };
    if (!validRangeBounds(bounds)) return { kind: 'error', code: 15, text: '#VALUE!' };
    return { kind: 'number', value: fn === 'SHEET' ? sheet + 1 : 1 };
  };
  return {
    matchExactRange,
    xmatchRange,
    indexRange,
    offsetValue,
    indirectValue,
    tableLookup,
    xlookupRange,
    vectorLookup,
    cellInfo,
    sheetInfo,
  };
}
