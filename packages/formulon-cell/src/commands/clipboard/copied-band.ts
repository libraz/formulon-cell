import { addrKey, MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { addMergeToMaps } from '../../store/merge-maps.js';
import type { CellFormat, SpreadsheetStore, State } from '../../store/store.js';
import { copy } from './copy.js';
import { type ClipboardSnapshot, captureSnapshotFromCopyResult } from './snapshot.js';

export function cloneCellFormat(format: CellFormat): CellFormat {
  return {
    ...format,
    borders: format.borders ? { ...format.borders } : undefined,
  };
}

export function makeBand(
  targetSheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): Range {
  return wholeRows
    ? {
        sheet: targetSheet,
        r0: insertionAt,
        c0: 0,
        r1: insertionAt + count - 1,
        c1: MAX_COL,
      }
    : {
        sheet: targetSheet,
        r0: 0,
        c0: insertionAt,
        r1: MAX_ROW,
        c1: insertionAt + count - 1,
      };
}

export function shiftedRangeForInsert(
  range: Range,
  targetSheet: number,
  axis: 'row' | 'col',
  split: number,
  count: number,
): Range {
  if (range.sheet !== targetSheet) return { ...range };
  if (axis === 'row') {
    return {
      ...range,
      r0: range.r0 >= split ? range.r0 + count : range.r0,
      r1: range.r1 >= split ? range.r1 + count : range.r1,
    };
  }
  return {
    ...range,
    c0: range.c0 >= split ? range.c0 + count : range.c0,
    c1: range.c1 >= split ? range.c1 + count : range.c1,
  };
}

export function commentsFitAfterInsert(
  wb: WorkbookHandle,
  sheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): boolean {
  const max = wholeRows ? MAX_ROW : MAX_COL;
  const lastMovable = max - count;
  for (const comment of wb.getComments(sheet)) {
    const axis = wholeRows ? comment.row : comment.col;
    if (axis >= insertionAt && axis > lastMovable) return false;
  }
  return true;
}

export function materializedStateForSheet(
  state: State,
  wb: WorkbookHandle,
  sheet: number,
  active: Addr,
  selectionRange: Range,
): State {
  const cells = new Map(
    Array.from(wb.cells(sheet), (cell) => [
      addrKey(cell.addr),
      { value: cell.value, formula: cell.formula },
    ]),
  );
  return {
    ...state,
    data: { ...state.data, sheetIndex: sheet, cells },
    selection: {
      ...state.selection,
      active,
      anchor: active,
      range: selectionRange,
      extraRanges: [],
    },
  };
}

export function refreshCopiedBandSnapshot(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot,
  targetSheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
  mode: 'copy' | 'cut' = 'copy',
): ClipboardSnapshot {
  const state = store.getState();
  const logicalBefore = snapshot.logicalRange ?? snapshot.range;
  const axis = wholeRows ? 'row' : 'col';
  const uiRange = mode === 'copy' && state.ui.copyMode === 'copy' ? state.ui.copyRange : null;
  const logical =
    uiRange && uiRange.sheet === snapshot.range.sheet
      ? uiRange
      : mode === 'copy'
        ? shiftedRangeForInsert(logicalBefore, targetSheet, axis, insertionAt, count)
        : logicalBefore;
  const sourceState = materializedSourceState(state, wb, snapshot, logical);
  const sourceCopy = copy(sourceState);
  if (!sourceCopy)
    return shiftedSnapshotFallback(snapshot, logical, targetSheet, wholeRows, insertionAt, count);
  const refreshed = captureSnapshotFromCopyResult(sourceState, sourceCopy, mode);
  if (!refreshed)
    return shiftedSnapshotFallback(snapshot, logical, targetSheet, wholeRows, insertionAt, count);
  // A cross-sheet insertion cannot change the source sheet's topology or
  // dimensions. The active store slice belongs to the destination sheet, so
  // preserve the source snapshot's merge and layout metadata rather than
  // accidentally recapturing the destination sheet's sparse maps.
  if (snapshot.range.sheet !== targetSheet) {
    return {
      ...refreshed,
      merges: snapshot.merges,
      rowHeights: snapshot.rowHeights,
      colWidths: snapshot.colWidths,
    };
  }
  return refreshed;
}

function shiftedSnapshotFallback(
  snapshot: ClipboardSnapshot,
  logical: Range,
  targetSheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): ClipboardSnapshot {
  const axis = wholeRows ? 'row' : 'col';
  const movedSource =
    snapshot.range.sheet === targetSheet
      ? shiftedRangeForInsert(snapshot.range, targetSheet, axis, insertionAt, count)
      : snapshot.range;
  return {
    ...snapshot,
    range: movedSource,
    logicalRange: logical,
  };
}

function materializedSourceState(
  state: State,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot,
  logical: Range,
): State {
  const sourceSheet = snapshot.range.sheet;
  const cells = new Map(
    Array.from(wb.cells(sourceSheet), (cell) => [
      addrKey(cell.addr),
      { value: cell.value, formula: cell.formula },
    ]),
  );
  if (sourceSheet === state.data.sheetIndex) {
    return {
      ...state,
      data: { ...state.data, sheetIndex: sourceSheet, cells },
      selection: {
        ...state.selection,
        active: { sheet: sourceSheet, row: logical.r0, col: logical.c0 },
        anchor: { sheet: sourceSheet, row: logical.r0, col: logical.c0 },
        range: logical,
        extraRanges: [],
      },
    };
  }

  const formats = new Map<string, CellFormat>();
  for (let r = 0; r < snapshot.rows; r += 1) {
    for (let c = 0; c < snapshot.cols; c += 1) {
      const format = snapshot.cells[r]?.[c]?.format;
      if (format) {
        formats.set(
          addrKey({ sheet: sourceSheet, row: snapshot.range.r0 + r, col: snapshot.range.c0 + c }),
          format,
        );
      }
    }
  }
  const byAnchor = new Map<string, Range>();
  const byCell = new Map<string, string>();
  for (const merge of snapshot.merges ?? []) {
    if (
      merge.r0 < 0 ||
      merge.c0 < 0 ||
      merge.r1 < merge.r0 ||
      merge.c1 < merge.c0 ||
      merge.r1 >= snapshot.rows ||
      merge.c1 >= snapshot.cols
    ) {
      continue;
    }
    addMergeToMaps(byAnchor, byCell, {
      sheet: sourceSheet,
      r0: snapshot.range.r0 + merge.r0,
      c0: snapshot.range.c0 + merge.c0,
      r1: snapshot.range.r0 + merge.r1,
      c1: snapshot.range.c0 + merge.c1,
    });
  }
  const rowHeights = new Map<number, number>();
  for (const [offset, size] of snapshot.rowHeights ?? []) {
    if (Number.isInteger(offset) && offset >= 0 && Number.isFinite(size) && size > 0) {
      rowHeights.set(logical.r0 + offset, size);
    }
  }
  const colWidths = new Map<number, number>();
  for (const [offset, size] of snapshot.colWidths ?? []) {
    if (Number.isInteger(offset) && offset >= 0 && Number.isFinite(size) && size > 0) {
      colWidths.set(logical.c0 + offset, size);
    }
  }
  return {
    ...state,
    data: { ...state.data, sheetIndex: sourceSheet, cells },
    format: { ...state.format, formats },
    merges: { byAnchor, byCell },
    layout: {
      ...state.layout,
      rowHeights,
      colWidths,
    },
    selection: {
      ...state.selection,
      active: { sheet: sourceSheet, row: logical.r0, col: logical.c0 },
      anchor: { sheet: sourceSheet, row: logical.r0, col: logical.c0 },
      range: logical,
      extraRanges: [],
    },
  };
}
