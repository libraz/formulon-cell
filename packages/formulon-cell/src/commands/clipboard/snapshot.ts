import { addrKey } from '../../engine/address.js';
import type { CellValue, Range } from '../../engine/types.js';
import type { CellFormat, State } from '../../store/store.js';

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
  return { range, rows, cols, cells: grid, merges, mode };
}
