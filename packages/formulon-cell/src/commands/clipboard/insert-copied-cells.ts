import { addrKey, MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { rangeContainsRange, rangesIntersect, sameRange } from '../../store/selection-geometry.js';
import { mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import { canShiftMerges, shiftCells, shiftFormats, shiftMerges } from '../cell-shift.js';
import { coerceInputForCell, writeCoerced } from '../coerce-input.js';
import type { History } from '../history.js';
import { blockedByProtection, isCellWritable, isSheetProtected } from '../protection.js';
import { insertCols, insertRows } from '../structure.js';
import {
  cloneCellFormat,
  commentsFitAfterInsert,
  isWholeColumnRange,
  isWholeRowRange,
  makeBand,
  materializedStateForSheet,
  refreshCopiedBandSnapshot,
  shiftedRangeForInsert,
} from './copied-band.js';
import {
  finalStartForCut,
  insertCutCopiedBand,
  unsupportedMetadataBlocksBand,
} from './cut-band-move.js';
import { copySnapshotMerges, writeSnapshotIntoInsertedRange } from './insert-shift.js';
import { pasteSpecial, resolvePasteDestination } from './paste-special.js';
import type { ClipboardSnapshot } from './snapshot.js';
import { parseTSV } from './tsv.js';

export type InsertCopiedCellsDirection = 'right' | 'down';

export interface InsertCopiedCellsResult {
  writtenRange: Range;
}

/** Optional target for a whole-row/whole-column copied-band insertion. When
 * omitted, the active selection's top-left cell is used. */
export type InsertCopiedBandTarget = Range | Addr;

const MAX_INSERT_COPIED_CELLS = 100_000;

export function insertCopiedCellsFromTSV(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  text: string,
  direction: InsertCopiedCellsDirection,
  snapshot?: ClipboardSnapshot | null,
): InsertCopiedCellsResult | null {
  if (!text && !snapshot) return null;
  const rows = snapshot ? [] : parseTSV(text);
  if (!snapshot && rows.length === 0) return null;
  const height = snapshot?.rows ?? rows.length;
  const width = snapshot?.cols ?? rows.reduce((max, row) => Math.max(max, row.length), 0);
  if (width <= 0) return null;
  if (height * width > MAX_INSERT_COPIED_CELLS) return null;

  const state = store.getState();
  const origin = state.selection.active;
  const sheet = origin.sheet;
  if (blockedByProtection(store, sheet, 'insert copied cells')) return null;

  const affected: Range =
    direction === 'down'
      ? {
          sheet,
          r0: origin.row,
          c0: origin.col,
          r1: MAX_ROW,
          c1: Math.min(MAX_COL, origin.col + width - 1),
        }
      : {
          sheet,
          r0: origin.row,
          c0: origin.col,
          r1: Math.min(MAX_ROW, origin.row + height - 1),
          c1: MAX_COL,
        };
  if (!canShiftMerges(state, affected, direction)) {
    // eslint-disable-next-line no-console
    console.warn('formulon-cell: insert copied cells blocked — merge would be split');
    return null;
  }

  if (history) history.begin();
  try {
    shiftCells(wb, affected, direction, direction === 'down' ? height : width);
    shiftFormats(store, history, affected, direction, direction === 'down' ? height : width);
    shiftMerges(store, wb, history, affected, direction, direction === 'down' ? height : width);

    wb.withBatchedRecalc(() => {
      if (snapshot) {
        writeSnapshotIntoInsertedRange(store, wb, history, snapshot, origin);
        return;
      }
      for (let r = 0; r < rows.length; r += 1) {
        const cells = rows[r] ?? [];
        for (let c = 0; c < cells.length; c += 1) {
          const addr: Addr = { sheet, row: origin.row + r, col: origin.col + c };
          if (!isCellWritable(store.getState(), addr)) continue;
          writeCoerced(wb, addr, coerceInputForCell(store.getState(), addr, cells[c] ?? ''));
        }
      }
    });
    copySnapshotMerges(store, wb, history, origin, snapshot);
    wb.recalcAuto();
  } finally {
    if (history) history.end();
  }

  return {
    writtenRange: {
      sheet,
      r0: origin.row,
      c0: origin.col,
      r1: origin.row + height - 1,
      c1: origin.col + width - 1,
    },
  };
}

/**
 * Insert a copied whole-row or whole-column band at the active target. The
 * snapshot keeps the logical full-band selection while its cell payload is a
 * bounded materialization, so this command inserts the structural rows/cols
 * first and then pastes the compact payload into the newly-created band.
 *
 * The source is re-snapshotted from the workbook after the structural edit.
 * This matters when the insertion is on the same sheet and moves the source:
 * native insertion has already rewritten the source formulas and the copied
 * formula must be based on that live source position.
 */
export function insertCopiedBand(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  snapshot: ClipboardSnapshot | null | undefined,
  target?: InsertCopiedBandTarget,
): InsertCopiedCellsResult | null {
  if (!snapshot || (snapshot.mode !== 'copy' && snapshot.mode !== 'cut')) return null;
  const state = store.getState();
  const prepared = prepareBandSnapshot(store, wb, snapshot);
  if (!prepared) return null;
  snapshot = prepared;
  const logical = snapshot.logicalRange ?? snapshot.range;
  const wholeRows = isWholeRowRange(logical);
  const wholeCols = isWholeColumnRange(logical);
  if (wholeRows === wholeCols) return null;
  if (!validCopiedBandSnapshot(snapshot, logical)) return null;

  const targetAddr = targetAddress(state, target);
  if (!targetAddr || targetAddr.sheet !== state.data.sheetIndex) return null;
  if (isSheetProtected(state, targetAddr.sheet)) return null;

  const count = wholeRows ? logical.r1 - logical.r0 + 1 : logical.c1 - logical.c0 + 1;
  if (!Number.isInteger(count) || count <= 0) return null;
  const max = wholeRows ? MAX_ROW : MAX_COL;
  const insertionAt = wholeRows ? targetAddr.row : targetAddr.col;
  const targetSheet = targetAddr.sheet;
  if (!sourceAxisMatchesTarget(snapshot, wholeRows, targetSheet)) return null;
  // The structural primitive reserves one terminal index as the insertion
  // boundary (`normalizeAxisEdit` uses `max - at` as its capacity). Match it
  // here so a multi-row/column request cannot be silently clamped to a
  // smaller insert and then pasted into a larger reported band.
  if (insertionAt < 0 || insertionAt > max) return null;
  if (snapshot.mode === 'cut' && snapshot.range.sheet === targetSheet) {
    const sourceStart = wholeRows ? logical.r0 : logical.c0;
    const finalStart = finalStartForCut(sourceStart, count, insertionAt);
    if (finalStart < 0 || finalStart + count > max) return null;
  } else if (insertionAt + count > max) {
    return null;
  }
  // Excel refuses an Insert Copied Cells operation when the insertion point
  // falls inside the copied band. The source would otherwise be widened by
  // the structural edit and the copy marquee would no longer describe the
  // original payload. Insertion at the band's start (before it), or after
  // its end, remains valid.
  if (
    snapshot.range.sheet === targetSheet &&
    insertionAt >= (wholeRows ? logical.r0 : logical.c0) &&
    insertionAt <= (wholeRows ? logical.r1 : logical.c1) &&
    (snapshot.mode === 'cut' || insertionAt > (wholeRows ? logical.r0 : logical.c0))
  ) {
    return null;
  }
  const adjacentCutNoOp =
    snapshot.mode === 'cut' &&
    snapshot.range.sheet === targetSheet &&
    insertionAt === (wholeRows ? logical.r0 : logical.c0) + count;
  if (
    snapshot.mode === 'cut' &&
    !adjacentCutNoOp &&
    unsupportedMetadataBlocksBand(store.getState(), wb, snapshot.range.sheet, targetSheet)
  ) {
    return null;
  }
  // A cut spans delete + insert + paste. Without a history frame there is no
  // command-local inverse to make a later native failure atomic, so refuse the
  // mutating form when callers did not provide the workbook history journal.
  if (snapshot.mode === 'cut' && !adjacentCutNoOp && history === null) return null;
  if (snapshot.rows <= 0 || snapshot.cols <= 0) return null;
  if (snapshot.rows * snapshot.cols > MAX_INSERT_COPIED_CELLS) return null;

  if (snapshot.mode === 'cut') {
    const cutStart =
      snapshot.range.sheet === targetSheet
        ? finalStartForCut(wholeRows ? logical.r0 : logical.c0, count, insertionAt)
        : insertionAt;
    return insertCutCopiedBand(
      store,
      wb,
      history,
      snapshot,
      targetAddr,
      logical,
      wholeRows,
      insertionAt,
      count,
      makeBand(targetSheet, wholeRows, cutStart, count),
    );
  }

  const band: Range = wholeRows
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
  const payloadDestination: Range = wholeRows
    ? {
        sheet: targetSheet,
        r0: insertionAt,
        c0: snapshot.range.c0,
        r1: insertionAt + snapshot.rows - 1,
        c1: snapshot.range.c1,
      }
    : {
        sheet: targetSheet,
        r0: snapshot.range.r0,
        c0: insertionAt,
        r1: snapshot.range.r1,
        c1: insertionAt + snapshot.cols - 1,
      };
  if (
    payloadDestination.r0 < 0 ||
    payloadDestination.c0 < 0 ||
    payloadDestination.r1 > MAX_ROW ||
    payloadDestination.c1 > MAX_COL
  ) {
    return null;
  }

  // Resolve the destination and validate merges before inserting anything.
  // The structural primitive is deliberately the first mutating operation;
  // a failed paste after it would leave a blank row/column behind.  Existing
  // destination merges are shifted by the same axis edit for this check, so
  // a merge crossing the inserted band's perpendicular edge is rejected in
  // the same way as pasteSpecial's matrix preflight.
  if (
    !preflightCopiedBand(
      store.getState(),
      wb,
      snapshot,
      band,
      targetSheet,
      wholeRows,
      insertionAt,
      count,
    )
  ) {
    return null;
  }
  if (!commentsFitAfterInsert(wb, targetSheet, wholeRows, insertionAt, count)) return null;

  // The structural primitive validates protection, worksheet overflow, and
  // the native operation result. The cheap checks above keep failures before
  // the mutation path; the boolean result prevents a paste after a failed
  // native insert.
  if (history) history.begin();
  try {
    const inserted = wholeRows
      ? insertRows(store, wb, history, insertionAt, count)
      : insertCols(store, wb, history, insertionAt, count);
    if (!inserted) return null;

    const refreshed = refreshCopiedBandSnapshot(
      store,
      wb,
      snapshot,
      targetSheet,
      wholeRows,
      insertionAt,
      count,
    );
    const pasteOrigin: Addr = wholeRows
      ? { sheet: targetSheet, row: insertionAt, col: 0 }
      : { sheet: targetSheet, row: 0, col: insertionAt };
    mutators.setActive(store, pasteOrigin);
    mutators.setRange(store, band);
    const pasteState = materializedStateForSheet(
      store.getState(),
      wb,
      targetSheet,
      pasteOrigin,
      band,
    );
    // `preflightCopiedBand` checked this geometry before the structural edit.
    // Keep the guard for defensive callers that mutate the store from a
    // subscription while the native operation is running; it cannot be a
    // normal failure path after the preflight above.
    const destination = resolvePasteDestination(pasteState, refreshed, false);
    if (!destination || !sameRange(destination, band))
      throw new Error('formulon-cell: copied band destination changed during insert');
    const pasted = pasteSpecial(
      pasteState,
      store,
      wb,
      refreshed,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    if (!pasted) throw new Error('formulon-cell: copied band paste failed after structural insert');
    mutators.setRange(store, band);
    return { writtenRange: band };
  } finally {
    if (history) history.end();
  }
}

function sourceAxisMatchesTarget(
  snapshot: ClipboardSnapshot,
  wholeRows: boolean,
  targetSheet: number,
): boolean {
  if (snapshot.range.sheet !== targetSheet) return true;
  const logical = snapshot.logicalRange ?? snapshot.range;
  return wholeRows
    ? logical.c0 === 0 && logical.c1 >= MAX_COL
    : logical.r0 === 0 && logical.r1 >= MAX_ROW;
}

/**
 * Rebuild the bounded payload from the engine before any structural edit.
 * The UI snapshot normally comes from the store, while notes and cells loaded
 * directly by the native workbook may not have reached that cache yet. A
 * whole-band cut must include those cells even when they sit outside the
 * store-derived compact range, otherwise delete-first would discard them.
 */
function prepareBandSnapshot(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot,
): ClipboardSnapshot | null {
  const logical = snapshot.logicalRange ?? snapshot.range;
  const wholeRows = isWholeRowRange(logical);
  const wholeCols = isWholeColumnRange(logical);
  const sourceBefore = snapshot.range;
  const comments = wb
    .getComments(sourceBefore.sheet)
    .filter((comment) =>
      wholeRows
        ? comment.row >= logical.r0 && comment.row <= logical.r1
        : wholeCols
          ? comment.col >= logical.c0 && comment.col <= logical.c1
          : comment.row >= sourceBefore.r0 &&
            comment.row <= sourceBefore.r1 &&
            comment.col >= sourceBefore.c0 &&
            comment.col <= sourceBefore.c1,
    );

  // Keep malformed caller snapshots malformed so the normal validation gate
  // rejects them before a structural mutation. Only a well-shaped payload may
  // be enriched with native notes.
  const expectedRows = sourceBefore.r1 - sourceBefore.r0 + 1;
  const expectedCols = sourceBefore.c1 - sourceBefore.c0 + 1;
  if (
    snapshot.rows !== expectedRows ||
    snapshot.cols !== expectedCols ||
    !Array.isArray(snapshot.cells) ||
    snapshot.cells.length !== expectedRows ||
    snapshot.cells.some((line) => !Array.isArray(line) || line.length !== expectedCols)
  ) {
    return snapshot;
  }
  if (comments.length === 0) return snapshot;

  const source = { ...sourceBefore };
  for (const comment of comments) {
    if (wholeRows) {
      source.c0 = Math.min(source.c0, comment.col);
      source.c1 = Math.max(source.c1, comment.col);
    } else if (wholeCols) {
      source.r0 = Math.min(source.r0, comment.row);
      source.r1 = Math.max(source.r1, comment.row);
    }
  }
  const rows = source.r1 - source.r0 + 1;
  const cols = source.c1 - source.c0 + 1;
  if (rows <= 0 || cols <= 0 || rows * cols > MAX_INSERT_COPIED_CELLS) return null;

  const state = store.getState();
  const engineCells = new Map(
    Array.from(wb.physicalCells(source.sheet), (cell) => [addrKey(cell.addr), cell]),
  );
  const commentByKey = new Map(
    comments.map((comment) => [
      addrKey({ sheet: source.sheet, row: comment.row, col: comment.col }),
      comment,
    ]),
  );
  const cells: ClipboardSnapshot['cells'] = [];
  for (let row = source.r0; row <= source.r1; row += 1) {
    const line: ClipboardSnapshot['cells'][number] = [];
    for (let col = source.c0; col <= source.c1; col += 1) {
      const sourceRow = row - sourceBefore.r0;
      const sourceCol = col - sourceBefore.c0;
      const original =
        sourceRow >= 0 &&
        sourceRow < snapshot.cells.length &&
        sourceCol >= 0 &&
        sourceCol < (snapshot.cells[sourceRow]?.length ?? 0)
          ? snapshot.cells[sourceRow]?.[sourceCol]
          : undefined;
      const key = addrKey({ sheet: source.sheet, row, col });
      const engineCell = engineCells.get(key);
      const storeFormat = state.format.formats.get(key);
      const cell = original
        ? {
            ...original,
            format: original.format
              ? cloneCellFormat(original.format)
              : storeFormat
                ? cloneCellFormat(storeFormat)
                : undefined,
          }
        : {
            value: engineCell?.value ?? { kind: 'blank' as const },
            formula: engineCell?.formula ?? null,
            format: storeFormat ? cloneCellFormat(storeFormat) : undefined,
          };
      const comment = commentByKey.get(key);
      if (comment) {
        cell.format = {
          ...(cell.format ?? {}),
          comment: comment.text,
          commentAuthor: comment.author || undefined,
        };
      }
      line.push(cell);
    }
    cells.push(line);
  }

  const merges = (snapshot.merges ?? []).map((merge) => ({
    ...merge,
    r0: sourceBefore.r0 + merge.r0 - source.r0,
    c0: sourceBefore.c0 + merge.c0 - source.c0,
    r1: sourceBefore.r0 + merge.r1 - source.r0,
    c1: sourceBefore.c0 + merge.c1 - source.c0,
  }));
  return {
    ...snapshot,
    range: source,
    rows,
    cols,
    cells,
    merges,
  };
}

function targetAddress(state: State, target?: InsertCopiedBandTarget): Addr | null {
  if (!target) {
    const selected = state.selection.range;
    return {
      sheet: selected.sheet,
      row: Math.min(selected.r0, selected.r1),
      col: Math.min(selected.c0, selected.c1),
    };
  }
  if ('r0' in target) return { sheet: target.sheet, row: target.r0, col: target.c0 };
  return target;
}

function validCopiedBandSnapshot(snapshot: ClipboardSnapshot, logical: Range): boolean {
  const source = snapshot.range;
  const validRange = (range: Range): boolean =>
    Number.isInteger(range.sheet) &&
    Number.isInteger(range.r0) &&
    Number.isInteger(range.c0) &&
    Number.isInteger(range.r1) &&
    Number.isInteger(range.c1) &&
    range.r0 >= 0 &&
    range.c0 >= 0 &&
    range.r1 >= range.r0 &&
    range.c1 >= range.c0 &&
    range.r1 <= MAX_ROW &&
    range.c1 <= MAX_COL;
  if (!validRange(source) || !validRange(logical) || source.sheet !== logical.sheet) return false;
  if (!rangeContainsRange(logical, source)) return false;
  if (snapshot.rows !== source.r1 - source.r0 + 1) return false;
  if (snapshot.cols !== source.c1 - source.c0 + 1) return false;
  if (snapshot.rows <= 0 || snapshot.cols <= 0) return false;
  if (
    !Array.isArray(snapshot.cells) ||
    snapshot.cells.length !== snapshot.rows ||
    snapshot.cells.some((row) => !Array.isArray(row) || row.length !== snapshot.cols)
  ) {
    return false;
  }
  for (const merge of snapshot.merges ?? []) {
    if (
      !Number.isInteger(merge.r0) ||
      !Number.isInteger(merge.c0) ||
      !Number.isInteger(merge.r1) ||
      !Number.isInteger(merge.c1) ||
      merge.r0 < 0 ||
      merge.c0 < 0 ||
      merge.r1 < merge.r0 ||
      merge.c1 < merge.c0 ||
      merge.r1 >= snapshot.rows ||
      merge.c1 >= snapshot.cols ||
      (merge.sheet !== undefined && merge.sheet !== source.sheet)
    ) {
      return false;
    }
  }
  return true;
}

function preflightCopiedBand(
  state: State,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot,
  band: Range,
  targetSheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): boolean {
  const pasteOrigin: Addr = wholeRows
    ? { sheet: targetSheet, row: insertionAt, col: 0 }
    : { sheet: targetSheet, row: 0, col: insertionAt };
  const pasteState = materializedStateForSheet(state, wb, targetSheet, pasteOrigin, band);
  const destination = resolvePasteDestination(pasteState, snapshot, false);
  if (!destination || !sameRange(destination, band)) return false;

  // `insertRows`/`insertCols` widens a merge when the split falls inside it;
  // otherwise it translates the merge. Predict that topology before the
  // structural edit and reject a destination that would only partially cover
  // one of those rectangles.
  const axis = wholeRows ? 'row' : 'col';
  for (const merge of state.merges.byAnchor.values()) {
    const shifted = shiftedRangeForInsert(merge, targetSheet, axis, insertionAt, count);
    if (!rangesIntersect(shifted, band)) continue;
    if (!rangeContainsRange(band, shifted)) return false;
  }
  return true;
}
