/** Where an internal paste lands: tile repetition, whole-row/-column band destinations, and the bounded cells a band paste materializes. */
import type { Addr, Range } from '../../engine/types.js';
import type { State } from '../../store/store.js';
import type { ClipboardSnapshot } from './snapshot.js';

export const MAX_ROW = 1_048_575;
export const MAX_COL = 16_383;
export const MAX_PASTE_CELLS = 1_000_000;

type BandAxis = 'row' | 'column';

export interface MaterializedPasteCell {
  row: number;
  col: number;
  sourceRow: number;
  sourceCol: number;
  sourceRowIndex: number;
  sourceColIndex: number;
}

export const logicalRangeFor = (snap: ClipboardSnapshot): Range => snap.logicalRange ?? snap.range;

export const bandAxisFor = (range: Range): BandAxis | null => {
  if (range.r0 === 0 && range.r1 >= MAX_ROW) return 'column';
  if (range.c0 === 0 && range.c1 >= MAX_COL) return 'row';
  return null;
};

const normalizedSelection = (state: State): Range => {
  const selected = state.selection.range;
  return {
    sheet: selected.sheet,
    r0: Math.min(selected.r0, selected.r1),
    c0: Math.min(selected.c0, selected.c1),
    r1: Math.max(selected.r0, selected.r1),
    c1: Math.max(selected.c0, selected.c1),
  };
};

/** Expand a whole-row/-column source to the logical destination band. Excel
 * accepts a single anchor cell only on the first row/column of the sheet;
 * another row/column is outside the band and is rejected. */
const destinationBandFor = (
  state: State,
  snap: ClipboardSnapshot,
  transpose: boolean,
): { range: Range; axis: BandAxis } | null => {
  if (transpose) return null;
  const logical = logicalRangeFor(snap);
  const axis = bandAxisFor(logical);
  if (!axis) return null;
  const selected = normalizedSelection(state);
  if (selected.sheet !== state.selection.active.sheet) return null;
  if (axis === 'column') {
    const firstRowOnly = selected.r0 === 0 && selected.r1 === 0;
    const fullRows = selected.r0 === 0 && selected.r1 >= MAX_ROW;
    if (!firstRowOnly && !fullRows) return null;
    const selectedWidth = selected.c1 - selected.c0 + 1;
    const logicalWidth = logical.c1 - logical.c0 + 1;
    const width =
      snap.mode === 'copy' && selectedWidth >= logicalWidth && selectedWidth % logicalWidth === 0
        ? selectedWidth
        : logicalWidth;
    if (selected.c0 + width - 1 > MAX_COL) return null;
    return {
      axis,
      range: {
        sheet: selected.sheet,
        r0: 0,
        c0: selected.c0,
        r1: MAX_ROW,
        c1: selected.c0 + width - 1,
      },
    };
  }
  const firstColOnly = selected.c0 === 0 && selected.c1 === 0;
  const fullCols = selected.c0 === 0 && selected.c1 >= MAX_COL;
  if (!firstColOnly && !fullCols) return null;
  const selectedHeight = selected.r1 - selected.r0 + 1;
  const logicalHeight = logical.r1 - logical.r0 + 1;
  const height =
    snap.mode === 'copy' && selectedHeight >= logicalHeight && selectedHeight % logicalHeight === 0
      ? selectedHeight
      : logicalHeight;
  if (selected.r0 + height - 1 > MAX_ROW) return null;
  return {
    axis,
    range: {
      sheet: selected.sheet,
      r0: selected.r0,
      c0: 0,
      r1: selected.r0 + height - 1,
      c1: MAX_COL,
    },
  };
};

const bandMaterializedCellCount = (
  snap: ClipboardSnapshot,
  destination: Range,
  axis: BandAxis,
): number =>
  snap.rows *
  snap.cols *
  Math.ceil(
    (axis === 'column'
      ? destination.c1 - destination.c0 + 1
      : destination.r1 - destination.r0 + 1) /
      (axis === 'column'
        ? logicalRangeFor(snap).c1 - logicalRangeFor(snap).c0 + 1
        : logicalRangeFor(snap).r1 - logicalRangeFor(snap).r0 + 1),
  );

/** Enumerate only the bounded source payload projected onto a whole-band
 * destination. The full logical axis is intentionally never materialized. */
export function materializedPasteCells(
  snap: ClipboardSnapshot,
  destination: Range,
  transpose = false,
): MaterializedPasteCell[] | null {
  if (transpose) return null;
  const logical = logicalRangeFor(snap);
  const axis = bandAxisFor(logical);
  if (!axis || bandAxisFor(destination) !== axis) return null;
  if (bandMaterializedCellCount(snap, destination, axis) > MAX_PASTE_CELLS) return null;
  const out: MaterializedPasteCell[] = [];
  const seen = new Set<string>();
  const add = (row: number, col: number, sourceRowIndex: number, sourceColIndex: number): void => {
    if (
      row < destination.r0 ||
      row > destination.r1 ||
      col < destination.c0 ||
      col > destination.c1
    ) {
      return;
    }
    const key = `${row}:${col}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({
      row,
      col,
      sourceRow: snap.range.r0 + sourceRowIndex,
      sourceCol: snap.range.c0 + sourceColIndex,
      sourceRowIndex,
      sourceColIndex,
    });
  };
  if (axis === 'column') {
    const logicalWidth = logical.c1 - logical.c0 + 1;
    for (let baseCol = destination.c0; baseCol <= destination.c1; baseCol += logicalWidth) {
      for (let sr = 0; sr < snap.rows; sr += 1) {
        for (let sc = 0; sc < snap.cols; sc += 1) {
          add(
            destination.r0 + snap.range.r0 + sr - logical.r0,
            baseCol + snap.range.c0 + sc - logical.c0,
            sr,
            sc,
          );
        }
      }
    }
  } else {
    const logicalHeight = logical.r1 - logical.r0 + 1;
    for (let baseRow = destination.r0; baseRow <= destination.r1; baseRow += logicalHeight) {
      for (let sr = 0; sr < snap.rows; sr += 1) {
        for (let sc = 0; sc < snap.cols; sc += 1) {
          add(
            baseRow + snap.range.r0 + sr - logical.r0,
            destination.c0 + snap.range.c0 + sc - logical.c0,
            sr,
            sc,
          );
        }
      }
    }
  }
  return out;
}

const destinationRangeFor = (origin: Addr, rows: number, cols: number): Range | null => {
  if (rows <= 0 || cols <= 0) return null;
  const r1 = origin.row + rows - 1;
  const c1 = origin.col + cols - 1;
  if (origin.row < 0 || origin.col < 0 || r1 > MAX_ROW || c1 > MAX_COL) return null;
  return { sheet: origin.sheet, r0: origin.row, c0: origin.col, r1, c1 };
};

/**
 * Resolve the destination footprint for an internal clipboard paste.
 *
 * A copied matrix repeats only when the current primary selection is a
 * normalized, exact multiple of the (possibly transposed) source tile. This
 * includes the full logical footprint selected by a merged-cell click; the
 * paste path preserves that merge for a scalar source.
 */
export function resolvePasteDestination(
  state: State,
  snap: ClipboardSnapshot,
  transpose = false,
): Range | null {
  const tileRows = transpose ? snap.cols : snap.rows;
  const tileCols = transpose ? snap.rows : snap.cols;
  if (!Number.isInteger(tileRows) || !Number.isInteger(tileCols)) return null;
  if (tileRows <= 0 || tileCols <= 0) return null;

  const band = destinationBandFor(state, snap, transpose);
  if (bandAxisFor(logicalRangeFor(snap)) && !band) return null;
  if (band) {
    if (bandMaterializedCellCount(snap, band.range, band.axis) > MAX_PASTE_CELLS) return null;
    return band.range;
  }

  const active = state.selection.active;
  let origin = { row: active.row, col: active.col };
  let rows = tileRows;
  let cols = tileCols;
  const selected = state.selection.range;
  const normalized = {
    sheet: selected.sheet,
    r0: Math.min(selected.r0, selected.r1),
    c0: Math.min(selected.c0, selected.c1),
    r1: Math.max(selected.r0, selected.r1),
    c1: Math.max(selected.c0, selected.c1),
  };
  const selectedRows = normalized.r1 - normalized.r0 + 1;
  const selectedCols = normalized.c1 - normalized.c0 + 1;
  const canRepeat =
    snap.mode === 'copy' &&
    normalized.sheet === active.sheet &&
    normalized.r0 >= 0 &&
    normalized.c0 >= 0 &&
    selectedRows >= tileRows &&
    selectedCols >= tileCols &&
    selectedRows % tileRows === 0 &&
    selectedCols % tileCols === 0;
  if (canRepeat) {
    origin = { row: normalized.r0, col: normalized.c0 };
    rows = selectedRows;
    cols = selectedCols;
  }
  const destination = destinationRangeFor({ sheet: active.sheet, ...origin }, rows, cols);
  if (!destination || rows * cols > MAX_PASTE_CELLS) return null;
  return destination;
}
