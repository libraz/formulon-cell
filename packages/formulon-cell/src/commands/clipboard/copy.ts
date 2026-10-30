import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import { formatBoolean, formatEditableNumber, formatError } from '../../engine/value.js';
import { isDataTableFormula } from '../../engine/workbook-handle-data-tables.js';
import { isWholeColumnRange, isWholeRowRange, sameRange } from '../../store/selection-geometry.js';
import type { State } from '../../store/store.js';
import { encodeTSV } from './tsv.js';

export interface CopyResult {
  /** TSV payload. Formula cells emit their formula text verbatim (`=...`) so
   *  a paste into Excel / Sheets reconstructs the formula; plain cells emit
   *  their value. */
  tsv: string;
  /** The range that was copied. */
  range: Range;
  /** The original logical selection. For a single selection this is the same
   *  range as `range`; whole-row/column copies keep it while `range` is the
   *  logical full band and `payloadRanges` carries the bounded materialized
   *  cells. */
  logicalRange?: Range;
  /** All copied ranges when copying a disjoint selection. */
  ranges?: Range[];
  /** Materialized ranges used for the TSV payload. Whole-row/-column copies
   *  are trimmed to the used span so we do not create huge blank payloads. */
  payloadRanges?: Range[];
}

/**
 * Snapshot the current selection into a TSV payload. Cells come from the
 * store's cached cell map (no engine reads). Formula cells emit their formula
 * text verbatim so a paste into Excel / Sheets rebuilds the formula; plain
 * cells emit their cached value. When `locale` is supplied, scalar values use
 * the active profile's edit-input spellings; omitted locale keeps the
 * canonical English clipboard representation. `localizeFormula` affects only
 * the external TSV payload; structured snapshots capture canonical formulas.
 */
export function copy(
  state: State,
  locale?: string,
  localizeFormula?: (formula: string) => string,
): CopyResult | null {
  const normalized = normalizedCopyRanges(state);
  if (!normalized) return null;
  const { payloadRanges, sourceRanges: ranges } = normalized;
  const r = ranges[0];
  if (!r) return null;
  const sheet = state.data.sheetIndex;
  if (payloadRanges.some((range) => range.r1 < range.r0 || range.c1 < range.c0)) return null;
  // Cap the materialised payload; whole bands are already trimmed to their used span.
  const totalCells = payloadRanges.reduce(
    (sum, range) => sum + (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1),
    0,
  );
  if (totalCells > 1_000_000) return null;

  const grid: string[][] = [];
  for (const range of payloadRanges) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      const line: string[] = [];
      for (let col = range.c0; col <= range.c1; col += 1) {
        line.push(displayValue(state, sheet, row, col, locale, localizeFormula));
      }
      grid.push(line);
    }
  }
  return {
    tsv: encodeTSV(grid),
    range: r,
    logicalRange: { ...r },
    ...(ranges.length > 1 ? { ranges } : {}),
    ...(payloadRanges.length > 1 || !payloadRanges[0] || !sameRange(payloadRanges[0], r)
      ? { payloadRanges }
      : {}),
  };
}

function normalizedCopyRanges(
  state: State,
): { sourceRanges: Range[]; payloadRanges: Range[] } | null {
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const sheet = state.data.sheetIndex;
  if (ranges.some((r) => r.sheet !== sheet)) return null;
  const sourceRanges = ranges
    .map((r) => ({ ...r }))
    .sort((a, b) => (a.r0 === b.r0 ? a.c0 - b.c0 : a.r0 - b.r0));
  const payloadRanges = trimWholeBandsToUsedSpan(state, sourceRanges);
  if (payloadRanges.length === 1) return { sourceRanges, payloadRanges };
  const c0 = ranges[0]?.c0;
  const c1 = ranges[0]?.c1;
  if (c0 === undefined || c1 === undefined) return null;
  // Multi-range copy is only supported when every band shares the same
  // column span. This covers Ctrl/Cmd row-header selections and same-width
  // disjoint cell bands without fabricating ragged TSV rows.
  if (payloadRanges.some((r) => r.c0 !== payloadRanges[0]?.c0 || r.c1 !== payloadRanges[0]?.c1)) {
    return null;
  }
  return { sourceRanges, payloadRanges };
}

function trimWholeBandsToUsedSpan(state: State, ranges: Range[]): Range[] {
  const fullSheet = ranges.length === 1 ? ranges[0] : undefined;
  // Whole-sheet selection: trim to the bounding box of used cells, formats and merges.
  if (fullSheet && isWholeRowRange(fullSheet) && isWholeColumnRange(fullSheet)) {
    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = -1;
    let minCol = Number.POSITIVE_INFINITY;
    let maxCol = -1;
    const include = (row: number, col: number): void => {
      if (row < fullSheet.r0 || row > fullSheet.r1 || col < fullSheet.c0 || col > fullSheet.c1) {
        return;
      }
      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    };
    const visitKey = (key: string): void => {
      const addr = parseAddrKey(key);
      if (
        addr &&
        addrKey(addr) === key &&
        addr.sheet === state.data.sheetIndex &&
        addr.row >= 0 &&
        addr.row <= MAX_ROW &&
        addr.col >= 0 &&
        addr.col <= MAX_COL
      ) {
        include(addr.row, addr.col);
      }
    };
    for (const key of state.data.cells.keys()) visitKey(key);
    for (const [key, format] of state.format.formats) {
      if (Object.keys(format).length > 0) visitKey(key);
    }
    for (const merge of state.merges.byAnchor.values()) {
      if (
        merge.sheet === state.data.sheetIndex &&
        merge.r0 >= 0 &&
        merge.c0 >= 0 &&
        merge.r1 <= MAX_ROW &&
        merge.c1 <= MAX_COL &&
        merge.r0 >= fullSheet.r0 &&
        merge.c0 >= fullSheet.c0 &&
        merge.r1 <= fullSheet.r1 &&
        merge.c1 <= fullSheet.c1 &&
        merge.r1 >= merge.r0 &&
        merge.c1 >= merge.c0
      ) {
        include(merge.r0, merge.c0);
        include(merge.r1, merge.c1);
      }
    }
    if (maxRow < 0 || maxCol < 0) {
      minRow = 0;
      maxRow = 0;
      minCol = 0;
      maxCol = 0;
    }
    return [{ sheet: fullSheet.sheet, r0: minRow, c0: minCol, r1: maxRow, c1: maxCol }];
  }

  const wholeRows = ranges.every(isWholeRowRange);
  const wholeCols = ranges.every(isWholeColumnRange);
  if (!wholeRows && !wholeCols) return ranges;

  let min = Number.POSITIVE_INFINITY;
  let max = -1;
  const visitKey = (key: string): void => {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== state.data.sheetIndex) return;
    const { row, col } = addr;
    if (wholeRows && ranges.some((r) => row >= r.r0 && row <= r.r1)) {
      min = Math.min(min, col);
      max = Math.max(max, col);
    } else if (wholeCols && ranges.some((r) => col >= r.c0 && col <= r.c1)) {
      min = Math.min(min, row);
      max = Math.max(max, row);
    }
  };
  for (const key of state.data.cells.keys()) visitKey(key);
  for (const [key, fmt] of state.format.formats) {
    if (Object.keys(fmt).length === 0) continue;
    visitKey(key);
  }
  // A merge is meaningful even when its anchor has no value or direct
  // formatting. Keep its complete perpendicular span in the compact payload
  // so the structured snapshot can reproduce the topology on paste.
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== state.data.sheetIndex) continue;
    if (
      wholeRows &&
      ranges.some(
        (r) => merge.r0 >= r.r0 && merge.r1 <= r.r1 && merge.c1 >= r.c0 && merge.c0 <= r.c1,
      )
    ) {
      min = Math.min(min, merge.c0);
      max = Math.max(max, merge.c1);
    } else if (
      wholeCols &&
      ranges.some(
        (r) => merge.c0 >= r.c0 && merge.c1 <= r.c1 && merge.r1 >= r.r0 && merge.r0 <= r.r1,
      )
    ) {
      min = Math.min(min, merge.r0);
      max = Math.max(max, merge.r1);
    }
  }
  if (max < 0) {
    min = 0;
    max = 0;
  }
  return ranges.map((r) => (wholeRows ? { ...r, c0: min, c1: max } : { ...r, r0: min, r1: max }));
}

function displayValue(
  state: State,
  sheet: number,
  row: number,
  col: number,
  locale?: string,
  localizeFormula?: (formula: string) => string,
): string {
  const key = `${sheet}:${row}:${col}`;
  const cell = state.data.cells.get(key);
  const hyperlinkDisplay = state.format.formats.get(key)?.hyperlinkDisplay;
  if (!cell) return hyperlinkDisplay ?? '';
  // Formula cells emit the formula text (already `=`-prefixed) so a paste
  //  into Excel / Sheets reconstructs the formula instead of freezing the
  //  last computed value.
  // A data table's TABLE() cells copy as their values, as the desktop app does.
  if (cell.formula != null && !isDataTableFormula(cell.formula)) {
    return localizeFormula?.(cell.formula) ?? cell.formula;
  }
  switch (cell.value.kind) {
    case 'number':
      return formatEditableNumber(cell.value.value, locale);
    case 'text':
      return cell.value.value;
    case 'bool':
      return formatBoolean(cell.value.value, locale);
    case 'error':
      return formatError(cell.value.code, cell.value.text, locale);
    default:
      return hyperlinkDisplay ?? '';
  }
}
