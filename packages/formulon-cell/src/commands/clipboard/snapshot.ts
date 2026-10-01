import { addrKey } from '../../engine/address.js';
import type { CellValue, Range } from '../../engine/types.js';
import type { CellFormat, State } from '../../store/store.js';
import type { CopyResult } from './copy.js';

/**
 * Structured snapshot of a range — values, formulas, AND formats. Captured
 * on copy/cut so a subsequent Paste Special can pick which of the three
 * layers to apply. Spreadsheets keep this internal-clipboard separate from the
 * system clipboard; we mirror that.
 */
export interface ClipboardCell {
  formula: string | null;
  value: CellValue;
  format: CellFormat | undefined;
}

/**
 * A merged rectangle relative to the top-left corner of the captured range.
 * The optional `sheet` member is accepted for compatibility with callers that
 * build snapshots from ordinary `Range` records; captured snapshots leave it
 * out because all rectangles belong to `ClipboardSnapshot.range.sheet`.
 */
export interface ClipboardMerge {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  sheet?: number;
}

export interface ClipboardSnapshot {
  /** Original source range (sheet-relative). */
  range: Range;
  /** The full logical selection before a whole-row/column payload was
   *  compacted. Ordinary snapshots set this to `range`. */
  logicalRange?: Range;
  rows: number;
  cols: number;
  /** rows × cols matrix in row-major order. Empty source cells are present
   *  but with `value = { kind: 'blank' }` and undefined format. */
  cells: ClipboardCell[][];
  /** Merges wholly contained by `range`, expressed relative to `range.r0/c0`.
   *  Older extension callers may omit this member. */
  merges?: ClipboardMerge[];
  /** Whether the snapshot originated from a cut or a copy. Excel moves cut
   *  cells: their formulas are pasted verbatim (no relative re-anchoring),
   *  unlike a copy which shifts relative references by the paste offset. */
  mode: 'copy' | 'cut';
  /** Captured row heights keyed by offset from `logicalRange.r0`. These are
   *  populated for whole-row copies; ordinary snapshots leave them empty. */
  rowHeights?: Map<number, number>;
  /** Captured column widths keyed by offset from `logicalRange.c0`. These are
   *  populated for whole-column copies; ordinary snapshots leave them empty. */
  colWidths?: Map<number, number>;
}

const blankCell = (): ClipboardCell => ({
  formula: null,
  value: { kind: 'blank' },
  format: undefined,
});

export function captureSnapshot(
  state: State,
  range: Range,
  mode: 'copy' | 'cut' = 'copy',
  options: { ignorePartialMerges?: boolean } = {},
): ClipboardSnapshot | null {
  const rows = range.r1 - range.r0 + 1;
  const cols = range.c1 - range.c0 + 1;
  if (rows <= 0 || cols <= 0) return null;
  // Cap at ~1M cells like the TSV copy path.
  if (rows * cols > 1_000_000) return null;

  const sheet = range.sheet;
  const merges: ClipboardMerge[] = [];
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== sheet) continue;
    const intersects = !(
      merge.r1 < range.r0 ||
      merge.r0 > range.r1 ||
      merge.c1 < range.c0 ||
      merge.c0 > range.c1
    );
    if (!intersects) continue;
    // A partial merge cannot be represented by a relative rectangle. Refuse
    // the snapshot so a cut/copy never silently tears its source topology.
    if (merge.r0 < range.r0 || merge.c0 < range.c0 || merge.r1 > range.r1 || merge.c1 > range.c1) {
      if (options.ignorePartialMerges) continue;
      return null;
    }
    merges.push({
      r0: merge.r0 - range.r0,
      c0: merge.c0 - range.c0,
      r1: merge.r1 - range.r0,
      c1: merge.c1 - range.c0,
    });
  }
  merges.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0 || a.r1 - b.r1 || a.c1 - b.c1);

  const grid: ClipboardCell[][] = [];
  for (let r = 0; r < rows; r += 1) {
    const line: ClipboardCell[] = [];
    for (let c = 0; c < cols; c += 1) {
      const key = addrKey({ sheet, row: range.r0 + r, col: range.c0 + c });
      const cell = state.data.cells.get(key);
      const fmt = state.format.formats.get(key);
      if (!cell && !fmt) {
        line.push(blankCell());
        continue;
      }
      line.push({
        formula: cell?.formula ?? null,
        value: cell?.value ?? { kind: 'blank' },
        format: fmt ? { ...fmt, borders: fmt.borders ? { ...fmt.borders } : undefined } : undefined,
      });
    }
    grid.push(line);
  }
  return {
    range,
    logicalRange: { ...range },
    rows,
    cols,
    cells: grid,
    merges,
    mode,
    rowHeights: new Map(),
    colWidths: new Map(),
  };
}

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

const isWholeRowRange = (range: Range): boolean => range.c0 === 0 && range.c1 >= MAX_COL;
const isWholeColumnRange = (range: Range): boolean => range.r0 === 0 && range.r1 >= MAX_ROW;

/**
 * Capture the bounded, structured payload represented by `copy` while
 * retaining the original logical selection. This is the single entry point
 * used by UI clipboard paths; direct `captureSnapshot` remains useful for
 * ordinary bounded ranges and tests.
 */
export function captureSnapshotFromCopyResult(
  state: State,
  result: CopyResult,
  mode: 'copy' | 'cut' = 'copy',
): ClipboardSnapshot | null {
  const ranges = result.payloadRanges ?? [result.range];
  if (ranges.length !== 1 || !ranges[0]) return null;
  const payloadRange = ranges[0];
  const logicalRange = result.logicalRange ?? result.range;
  const logicalIsBand =
    (logicalRange.r0 === 0 && logicalRange.r1 >= MAX_ROW) ||
    (logicalRange.c0 === 0 && logicalRange.c1 >= MAX_COL);
  const snapshot = captureSnapshot(state, payloadRange, mode, {
    // Whole-band copies intentionally omit a merge crossing the selected
    // axis boundary. Excel copies the anchor/value but does not carry the
    // partial merge into the destination.
    ignorePartialMerges: logicalIsBand,
  });
  if (!snapshot) return null;
  const rowHeights = new Map<number, number>();
  const colWidths = new Map<number, number>();
  if (isWholeRowRange(logicalRange)) {
    for (const [row, height] of state.layout.rowHeights) {
      if (row >= logicalRange.r0 && row <= logicalRange.r1) {
        rowHeights.set(row - logicalRange.r0, height);
      }
    }
  }
  if (isWholeColumnRange(logicalRange)) {
    for (const [col, width] of state.layout.colWidths) {
      if (col >= logicalRange.c0 && col <= logicalRange.c1) {
        colWidths.set(col - logicalRange.c0, width);
      }
    }
  }
  return {
    ...snapshot,
    logicalRange: { ...logicalRange },
    rowHeights,
    colWidths,
  };
}
