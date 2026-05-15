import type { Range } from '../../engine/types.js';
import type { State } from '../../store/store.js';
import { encodeTSV } from './tsv.js';

export interface CopyResult {
  /** TSV payload — values only; formulas resolve to their displayed text. */
  tsv: string;
  /** The range that was copied. */
  range: Range;
  /** All copied ranges when copying a disjoint selection. */
  ranges?: Range[];
  /** Materialized ranges used for the TSV payload. Whole-row/-column copies
   *  are trimmed to the used span so we do not create huge blank payloads. */
  payloadRanges?: Range[];
}

/**
 * Snapshot the current selection into a TSV payload. Values come from the
 * store's cached cell map (no engine reads) — for formula cells this means
 * the last computed value, matching the spreadsheet's "values only" copy semantic.
 */
export function copy(state: State): CopyResult | null {
  const normalized = normalizedCopyRanges(state);
  if (!normalized) return null;
  const { payloadRanges, sourceRanges: ranges } = normalized;
  const r = ranges[0];
  if (!r) return null;
  const sheet = state.data.sheetIndex;
  if (payloadRanges.some((range) => range.r1 < range.r0 || range.c1 < range.c0)) return null;
  // Refuse to materialise an entire-sheet copy — 17B blank strings would OOM.
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
        line.push(displayValue(state, sheet, row, col));
      }
      grid.push(line);
    }
  }
  return {
    tsv: encodeTSV(grid),
    range: r,
    ...(ranges.length > 1 ? { ranges } : {}),
    ...(payloadRanges.length > 1 || !sameRange(payloadRanges[0], r) ? { payloadRanges } : {}),
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
  const wholeRows = ranges.every((r) => r.c0 === 0 && r.c1 >= 16383);
  const wholeCols = ranges.every((r) => r.r0 === 0 && r.r1 >= 1048575);
  if (!wholeRows && !wholeCols) return ranges;

  let min = Number.POSITIVE_INFINITY;
  let max = -1;
  for (const key of state.data.cells.keys()) {
    const [sheetRaw, rowRaw, colRaw] = key.split(':');
    const sheet = Number(sheetRaw);
    if (sheet !== state.data.sheetIndex) continue;
    const row = Number(rowRaw);
    const col = Number(colRaw);
    if (wholeRows && ranges.some((r) => row >= r.r0 && row <= r.r1)) {
      min = Math.min(min, col);
      max = Math.max(max, col);
    } else if (wholeCols && ranges.some((r) => col >= r.c0 && col <= r.c1)) {
      min = Math.min(min, row);
      max = Math.max(max, row);
    }
  }
  if (max < 0) {
    min = 0;
    max = 0;
  }
  return ranges.map((r) => (wholeRows ? { ...r, c0: min, c1: max } : { ...r, r0: min, r1: max }));
}

function sameRange(a: Range | undefined, b: Range | undefined): boolean {
  return (
    !!a &&
    !!b &&
    a.sheet === b.sheet &&
    a.r0 === b.r0 &&
    a.c0 === b.c0 &&
    a.r1 === b.r1 &&
    a.c1 === b.c1
  );
}

function displayValue(state: State, sheet: number, row: number, col: number): string {
  const cell = state.data.cells.get(`${sheet}:${row}:${col}`);
  if (!cell) return '';
  switch (cell.value.kind) {
    case 'number':
      return String(cell.value.value);
    case 'text':
      return cell.value.value;
    case 'bool':
      return cell.value.value ? 'TRUE' : 'FALSE';
    case 'error':
      return cell.value.text;
    default:
      return '';
  }
}
