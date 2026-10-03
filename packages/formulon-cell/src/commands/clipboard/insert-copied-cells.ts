import { addrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import { coerceInputForCell, writeCoerced } from '../coerce-input.js';
import { type AxisBandMoveContext, adjustFormulaForAxisBandMove } from '../formula-refs.js';
import { type History, recordMergesChangeWithEngine } from '../history.js';
import { isCellWritable, isSheetProtected } from '../protection.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../structure.js';
import { copy } from './copy.js';
import {
  axisVisibilityFromDimensions,
  captureCutTransientSnapshot,
  captureEngineXfSnapshot,
  captureSameSheetDimensionRestore,
  captureSourceDimensionOverrides,
  captureStoreAxisVisibility,
  consumeCutMarquee,
  emptyAxisVisibility,
  recordAxisVisibilityTransfer,
  recordSameSheetDimensionRestore,
  recordSourceFormatEngineRepair,
  resetCrossSheetSourceDimensions,
  restoreCutTransientSnapshot,
} from './cut-band-snapshots.js';
import {
  addMergeToMaps,
  type CellRecord,
  canShiftMerges,
  copySnapshotMerges,
  MAX_COL,
  MAX_ROW,
  rangesIntersect,
  removeIntersectingMerges,
  shiftCells,
  shiftFormats,
  shiftMerges,
  writeCell,
  writeSnapshotIntoInsertedRange,
} from './insert-shift.js';
import { pasteSpecial, resolvePasteDestination } from './paste-special.js';
import { type ClipboardSnapshot, captureSnapshotFromCopyResult } from './snapshot.js';
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
  if (isSheetProtected(state, sheet)) {
    // eslint-disable-next-line no-console
    console.warn(`formulon-cell: insert copied cells blocked — sheet ${sheet} is protected`);
    return null;
  }

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
  if (!canShiftMerges(state, affected, direction)) return null;

  if (history) history.begin();
  try {
    shiftCells(store.getState(), wb, affected, direction, direction === 'down' ? height : width);
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

const WHOLE_ROW_END = 1_048_575;
const WHOLE_COLUMN_END = 16_383;

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
  const max = wholeRows ? WHOLE_ROW_END : WHOLE_COLUMN_END;
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
        c1: WHOLE_COLUMN_END,
      }
    : {
        sheet: targetSheet,
        r0: 0,
        c0: insertionAt,
        r1: WHOLE_ROW_END,
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
    payloadDestination.r1 > WHOLE_ROW_END ||
    payloadDestination.c1 > WHOLE_COLUMN_END
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

function isWholeRowRange(range: Range): boolean {
  return range.c0 === 0 && range.c1 >= WHOLE_COLUMN_END;
}

function isWholeColumnRange(range: Range): boolean {
  return range.r0 === 0 && range.r1 >= WHOLE_ROW_END;
}

function sourceAxisMatchesTarget(
  snapshot: ClipboardSnapshot,
  wholeRows: boolean,
  targetSheet: number,
): boolean {
  if (snapshot.range.sheet !== targetSheet) return true;
  const logical = snapshot.logicalRange ?? snapshot.range;
  return wholeRows
    ? logical.c0 === 0 && logical.c1 >= WHOLE_COLUMN_END
    : logical.r0 === 0 && logical.r1 >= WHOLE_ROW_END;
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

function cloneCellFormat(format: CellFormat): CellFormat {
  return {
    ...format,
    borders: format.borders ? { ...format.borders } : undefined,
  };
}

function makeBand(
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
        c1: WHOLE_COLUMN_END,
      }
    : {
        sheet: targetSheet,
        r0: 0,
        c0: insertionAt,
        r1: WHOLE_ROW_END,
        c1: insertionAt + count - 1,
      };
}

function axisOfBand(wholeRows: boolean): 'row' | 'col' {
  return wholeRows ? 'row' : 'col';
}

function formulaLessSnapshot(snapshot: ClipboardSnapshot): ClipboardSnapshot {
  return {
    ...snapshot,
    mode: 'copy',
    cells: snapshot.cells.map((line) =>
      line.map((cell) => ({ ...cell, formula: null, format: cloneCellFormat(cell.format ?? {}) })),
    ),
  };
}

function finalStartForCut(sourceStart: number, count: number, insertionAt: number): number {
  return insertionAt < sourceStart ? insertionAt : insertionAt - count;
}

function mapCutAxisIndex(
  index: number,
  sourceStart: number,
  count: number,
  insertionAt: number,
): number {
  const finalStart = finalStartForCut(sourceStart, count, insertionAt);
  if (index >= sourceStart && index < sourceStart + count) {
    return finalStart + index - sourceStart;
  }
  const afterDelete = index >= sourceStart + count ? index - count : index;
  return afterDelete >= finalStart ? afterDelete + count : afterDelete;
}

function mapCutOwnerAddress(
  addr: Addr,
  targetSheet: number,
  axis: 'row' | 'col',
  sourceStart: number,
  count: number,
  insertionAt: number,
): Addr {
  if (addr.sheet !== targetSheet) return { ...addr };
  return axis === 'row'
    ? { ...addr, row: mapCutAxisIndex(addr.row, sourceStart, count, insertionAt) }
    : { ...addr, col: mapCutAxisIndex(addr.col, sourceStart, count, insertionAt) };
}

interface FormulaRecord {
  addr: Addr;
  formula: string;
}

function collectAllWorkbookFormulas(wb: WorkbookHandle): FormulaRecord[] {
  const out: FormulaRecord[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    for (const cell of wb.cells(sheet)) {
      if (cell.formula !== null) out.push({ addr: cell.addr, formula: cell.formula });
    }
  }
  return out;
}

function collectFormulaRecordsForAddresses(
  wb: WorkbookHandle,
  addresses: readonly Addr[],
): Map<string, CellRecord | null> {
  const wanted = new Set(addresses.map(addrKey));
  const records = new Map<string, CellRecord | null>();
  for (const addr of addresses) records.set(addrKey(addr), null);
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    for (const cell of wb.cells(sheet)) {
      const key = addrKey(cell.addr);
      if (!wanted.has(key)) continue;
      records.set(key, { addr: cell.addr, value: cell.value, formula: cell.formula });
    }
  }
  return records;
}

function applyFormulaOwnerRecords(
  wb: WorkbookHandle,
  records: ReadonlyMap<string, CellRecord | null>,
): void {
  wb.withBatchedRecalc(() => {
    for (const [key, record] of records) {
      const [sheetRaw, rowRaw, colRaw] = key.split(':');
      const addr = {
        sheet: Number(sheetRaw),
        row: Number(rowRaw),
        col: Number(colRaw),
      };
      if (!record) {
        wb.setBlank(addr);
      } else {
        writeCell(wb, addr, record.value, record.formula);
      }
    }
  });
}

function sameCellRecord(a: CellRecord | null, b: CellRecord | null): boolean {
  return (
    a?.formula === b?.formula &&
    a?.addr.sheet === b?.addr.sheet &&
    a?.addr.row === b?.addr.row &&
    a?.addr.col === b?.addr.col &&
    JSON.stringify(a?.value) === JSON.stringify(b?.value)
  );
}

function recordFormulaOwnerRewrite(
  wb: WorkbookHandle,
  history: History | null,
  original: readonly FormulaRecord[],
  axis: 'row' | 'col',
  targetSheet: number,
  sourceStart: number,
  count: number,
  insertionAt: number,
): void {
  const sheetNames = Array.from({ length: wb.sheetCount }, (_, sheet) => wb.sheetName(sheet));
  const contextFor = (formulaSheet: number): AxisBandMoveContext => ({
    axis,
    sheet: targetSheet,
    sourceStart,
    count,
    insertionAt,
    formulaSheet,
    sheetNames,
  });
  const mapped = original.map((entry) => ({
    ...entry,
    addr: mapCutOwnerAddress(entry.addr, targetSheet, axis, sourceStart, count, insertionAt),
    formula: adjustFormulaForAxisBandMove(entry.formula, contextFor(entry.addr.sheet)),
  }));
  const touched = new Map<string, Addr>();
  for (const entry of original) touched.set(addrKey(entry.addr), entry.addr);
  for (const entry of mapped) touched.set(addrKey(entry.addr), entry.addr);
  const before = collectFormulaRecordsForAddresses(wb, [...touched.values()]);
  wb.withBatchedRecalc(() => {
    for (const entry of mapped) wb.setFormula(entry.addr, entry.formula);
  });
  wb.recalcAuto();
  const after = collectFormulaRecordsForAddresses(wb, [...touched.values()]);
  if (!history) return;
  const changed = [...touched.keys()].some(
    (key) => !sameCellRecord(before.get(key) ?? null, after.get(key) ?? null),
  );
  if (!changed) return;
  history.push({
    undo: () => applyFormulaOwnerRecords(wb, before),
    redo: () => applyFormulaOwnerRecords(wb, after),
  });
}

function cutSourceMergesAreWhole(state: State, wb: WorkbookHandle, source: Range): boolean {
  const merges = new Map<string, Range>();
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet === source.sheet)
      merges.set(addrKey({ sheet: merge.sheet, row: merge.r0, col: merge.c0 }), merge);
  }
  for (const merge of wb.getMerges(source.sheet)) {
    merges.set(addrKey({ sheet: merge.sheet, row: merge.r0, col: merge.c0 }), merge);
  }
  for (const merge of merges.values()) {
    if (rangesIntersect(merge, source) && !rangeContains(source, merge)) return false;
  }
  return true;
}

function cutDestinationMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): void {
  const destination = makeBand(sheet, wholeRows, insertionAt, count);
  recordMergesChangeWithEngine(history, store, wb, sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map(s.merges.byAnchor);
      const byCell = new Map(s.merges.byCell);
      removeIntersectingMerges(byAnchor, byCell, destination);
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}

function cutBandFitsAfterReorder(
  state: State,
  wb: WorkbookHandle,
  source: Range,
  targetSheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): boolean {
  const sourceStart = wholeRows ? source.r0 : source.c0;
  const axis = axisOfBand(wholeRows);
  const max = wholeRows ? WHOLE_ROW_END : WHOLE_COLUMN_END;
  const finalStart = finalStartForCut(sourceStart, count, insertionAt);
  if (finalStart < 0 || finalStart + count > max) return false;
  const mappedIndex = (index: number): number =>
    mapCutAxisIndex(index, sourceStart, count, insertionAt);

  for (const cell of wb.cells(targetSheet)) {
    const index = axis === 'row' ? cell.addr.row : cell.addr.col;
    if (mappedIndex(index) > max) return false;
  }
  for (const key of state.format.formats.keys()) {
    const [sheetRaw, rowRaw, colRaw] = key.split(':');
    if (Number(sheetRaw) !== targetSheet) continue;
    const index = wholeRows ? Number(rowRaw) : Number(colRaw);
    if (mappedIndex(index) > max) return false;
  }
  const indexed = wholeRows
    ? [
        ...state.layout.rowHeights.keys(),
        ...state.layout.hiddenRows,
        ...state.layout.outlineRows.keys(),
      ]
    : [
        ...state.layout.colWidths.keys(),
        ...state.layout.hiddenCols,
        ...state.layout.outlineCols.keys(),
      ];
  if (indexed.some((index) => mappedIndex(index) > max)) return false;
  for (const comment of wb.getComments(targetSheet)) {
    if (mappedIndex(wholeRows ? comment.row : comment.col) > max) return false;
  }
  return true;
}

function cutBandPreflight(
  state: State,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot,
  targetAddr: Addr,
  logical: Range,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): boolean {
  const targetSheet = targetAddr.sheet;
  if (isSheetProtected(state, targetSheet) || isSheetProtected(state, snapshot.range.sheet)) {
    return false;
  }
  if (!cutSourceMergesAreWhole(state, wb, logical)) return false;
  if (snapshot.range.sheet === targetSheet) {
    if (!cutBandFitsAfterReorder(state, wb, logical, targetSheet, wholeRows, insertionAt, count)) {
      return false;
    }
  } else if (!commentsFitAfterInsert(wb, targetSheet, wholeRows, insertionAt, count)) {
    return false;
  }
  return true;
}

function hydrateSourceMergesInStore(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
): void {
  if (!wb.capabilities.merges) return;
  const sourceMerges = wb.getMerges(sheet);
  if (sourceMerges.length === 0) return;
  store.setState((s) => {
    const byAnchor = new Map(s.merges.byAnchor);
    const byCell = new Map(s.merges.byCell);
    for (const merge of sourceMerges) addMergeToMaps(byAnchor, byCell, merge);
    return { ...s, merges: { byAnchor, byCell } };
  });
}

interface SourceCommentSnapshot {
  addr: Addr;
  author: string;
  text: string;
  format: CellFormat | undefined;
}

function captureSourceCommentSnapshots(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  source: Range,
): SourceCommentSnapshot[] {
  const byKey = new Map<string, SourceCommentSnapshot>();
  for (const comment of wb.getComments(source.sheet)) {
    if (
      comment.row < source.r0 ||
      comment.row > source.r1 ||
      comment.col < source.c0 ||
      comment.col > source.c1
    ) {
      continue;
    }
    const addr = { sheet: source.sheet, row: comment.row, col: comment.col };
    byKey.set(addrKey(addr), {
      addr,
      author: comment.author,
      text: comment.text,
      format: undefined,
    });
  }
  for (const [key, format] of store.getState().format.formats) {
    const [sheetRaw, rowRaw, colRaw] = key.split(':');
    const addr = { sheet: Number(sheetRaw), row: Number(rowRaw), col: Number(colRaw) };
    if (
      addr.sheet !== source.sheet ||
      addr.row < source.r0 ||
      addr.row > source.r1 ||
      addr.col < source.c0 ||
      addr.col > source.c1 ||
      typeof format.comment !== 'string' ||
      format.comment.length === 0
    ) {
      continue;
    }
    const existing = byKey.get(key);
    byKey.set(key, {
      addr,
      author: format.commentAuthor ?? existing?.author ?? '',
      text: format.comment,
      format: cloneCellFormat(format),
    });
  }
  return [...byKey.values()];
}

function recordSourceCommentRestore(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  comments: readonly SourceCommentSnapshot[],
): void {
  if (!history || comments.length === 0) return;
  history.push({
    undo: () => {
      for (const comment of comments) {
        wb.setCommentEntry(
          comment.addr.sheet,
          comment.addr.row,
          comment.addr.col,
          comment.author,
          comment.text,
        );
      }
      store.setState((s) => {
        const formats = new Map(s.format.formats);
        for (const comment of comments) {
          const key = addrKey(comment.addr);
          const current = formats.get(key);
          formats.set(
            key,
            comment.format ?? {
              ...(current ?? {}),
              comment: comment.text,
              commentAuthor: comment.author || undefined,
            },
          );
        }
        return { ...s, format: { ...s.format, formats } };
      });
    },
    // The following delete operation removes the source comments again. No
    // forward mutation is needed here; this entry only repairs native delete
    // undo, which cannot reconstruct deleted comments by itself.
    redo: () => {},
  });
}

function insertCutCopiedBand(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  snapshot: ClipboardSnapshot,
  targetAddr: Addr,
  logical: Range,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
  band: Range,
): InsertCopiedCellsResult | null {
  const targetSheet = targetAddr.sheet;
  if (
    !cutBandPreflight(
      store.getState(),
      wb,
      snapshot,
      targetAddr,
      logical,
      wholeRows,
      insertionAt,
      count,
    )
  ) {
    return null;
  }
  if (
    snapshot.range.sheet === targetSheet &&
    insertionAt === (wholeRows ? logical.r0 : logical.c0) + count
  ) {
    const sourceStart = wholeRows ? logical.r0 : logical.c0;
    const sourceOrigin: Addr = wholeRows
      ? { sheet: targetSheet, row: sourceStart, col: 0 }
      : { sheet: targetSheet, row: 0, col: sourceStart };
    consumeCutMarquee(store);
    mutators.setActive(store, sourceOrigin);
    mutators.setRange(store, logical);
    return { writtenRange: logical };
  }
  const axis = axisOfBand(wholeRows);
  const sourceStart = wholeRows ? logical.r0 : logical.c0;
  const finalStart = finalStartForCut(sourceStart, count, insertionAt);
  const finalBand = makeBand(targetSheet, wholeRows, finalStart, count);
  const originalFormulas =
    snapshot.range.sheet === targetSheet ? collectAllWorkbookFormulas(wb) : [];
  const sourceComments =
    snapshot.range.sheet === targetSheet ? captureSourceCommentSnapshots(store, wb, logical) : [];
  const sourceDimensions =
    snapshot.range.sheet === targetSheet
      ? captureSameSheetDimensionRestore(store, wb, logical, wholeRows)
      : null;
  const sourceVisibility = axisVisibilityFromDimensions(
    captureSourceDimensionOverrides(wb, logical, wholeRows),
    snapshot.range.sheet === targetSheet
      ? captureStoreAxisVisibility(store, logical, wholeRows)
      : emptyAxisVisibility(),
  );
  const payload = formulaLessSnapshot(snapshot);
  const sourceEngineXfs = captureEngineXfSnapshot(wb, snapshot.range.sheet);
  const transientBefore = captureCutTransientSnapshot(store.getState());

  const transaction = history?.begin();
  let success = false;
  let aborted = false;
  let settled = false;
  let abortError: unknown;
  const abortTransaction = (): void => {
    if (transaction === undefined || settled || aborted) return;
    // Mark the frame closed before invoking user-supplied undo callbacks. An
    // undo callback may throw; the finally block must never attempt this token
    // a second time after History.abort has already removed the frame.
    aborted = true;
    settled = true;
    try {
      history?.abort(transaction);
    } catch (error) {
      abortError = error;
    }
  };
  try {
    recordSourceFormatEngineRepair(store, wb, history, snapshot.range.sheet, sourceEngineXfs, {
      wholeRows,
      ranges: snapshot.range.sheet === targetSheet ? [logical, finalBand] : [logical],
      ...(snapshot.range.sheet === targetSheet ? { extraIndex: sourceStart + count } : {}),
    });
    // A destination merge crossed by a cut insertion is removed in Excel. Do
    // this before either axis operation so the native engine cannot widen it
    // across the new band and leave an unrecoverable partial topology.
    cutDestinationMerges(store, wb, history, targetSheet, wholeRows, insertionAt, count);

    if (snapshot.range.sheet === targetSheet) {
      recordSourceCommentRestore(store, wb, history, sourceComments);
      if (sourceDimensions) {
        recordSameSheetDimensionRestore(store, wb, history, logical, wholeRows, sourceDimensions);
      }
      const deleted = wholeRows
        ? deleteRows(store, wb, history, sourceStart, count)
        : deleteCols(store, wb, history, sourceStart, count);
      if (!deleted) return null;
      const inserted = wholeRows
        ? insertRows(store, wb, history, finalStart, count)
        : insertCols(store, wb, history, finalStart, count);
      if (!inserted) return null;

      const pasteOrigin: Addr = wholeRows
        ? { sheet: targetSheet, row: finalStart, col: 0 }
        : { sheet: targetSheet, row: 0, col: finalStart };
      mutators.setActive(store, pasteOrigin);
      mutators.setRange(store, finalBand);
      const pasteState = materializedStateForSheet(
        store.getState(),
        wb,
        targetSheet,
        pasteOrigin,
        finalBand,
      );
      const pasted = pasteSpecial(
        pasteState,
        store,
        wb,
        payload,
        { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
        history,
      );
      if (!pasted) throw new Error('formulon-cell: cut band paste failed after structural edit');
      recordFormulaOwnerRewrite(
        wb,
        history,
        originalFormulas,
        axis,
        targetSheet,
        sourceStart,
        count,
        insertionAt,
      );
      recordAxisVisibilityTransfer(
        store,
        wb,
        history,
        logical,
        finalBand,
        wholeRows,
        sourceVisibility,
      );
      consumeCutMarquee(store);
      mutators.setActive(store, pasteOrigin);
      mutators.setRange(store, finalBand);
      success = true;
      return { writtenRange: finalBand };
    }

    // Cross-sheet cut keeps the source axis in place. Insert the destination
    // band first, then let Paste All clear the original source cells while
    // retaining the cut formula/reference semantics.
    const inserted = wholeRows
      ? insertRows(store, wb, history, insertionAt, count)
      : insertCols(store, wb, history, insertionAt, count);
    if (!inserted) return null;
    hydrateSourceMergesInStore(store, wb, snapshot.range.sheet);
    const refreshed = refreshCopiedBandSnapshot(
      store,
      wb,
      snapshot,
      targetSheet,
      wholeRows,
      insertionAt,
      count,
      'cut',
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
    const pasted = pasteSpecial(
      pasteState,
      store,
      wb,
      refreshed,
      { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
      history,
    );
    if (!pasted) throw new Error('formulon-cell: cross-sheet cut band paste failed');
    recordAxisVisibilityTransfer(store, wb, history, logical, band, wholeRows, sourceVisibility);
    resetCrossSheetSourceDimensions(store, wb, history, snapshot, wholeRows);
    mutators.setActive(store, pasteOrigin);
    mutators.setRange(store, band);
    success = true;
    return { writtenRange: band };
  } catch (error) {
    abortTransaction();
    restoreCutTransientSnapshot(store, transientBefore);
    if (abortError !== undefined) throw abortError;
    // eslint-disable-next-line no-console
    console.warn('formulon-cell: cut copied band failed', error);
    return null;
  } finally {
    try {
      if (transaction !== undefined && !settled) {
        if (success) {
          settled = true;
          history?.end(transaction);
        } else {
          abortTransaction();
        }
      }
    } catch (error) {
      success = false;
      restoreCutTransientSnapshot(store, transientBefore);
      // The operation has already returned its result by the time a history
      // commit runs in finally. Keep the transient UI safe even if a history
      // listener rejects the commit; the workbook mutation itself remains
      // covered by the transaction's normal error path.
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cut copied band history commit failed', error);
    }
    if (!success) {
      restoreCutTransientSnapshot(store, transientBefore);
    }
  }
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

function shiftedRangeForInsert(
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
    range.r1 <= WHOLE_ROW_END &&
    range.c1 <= WHOLE_COLUMN_END;
  if (!validRange(source) || !validRange(logical) || source.sheet !== logical.sheet) return false;
  if (!rangeContains(logical, source)) return false;
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
    if (!rangeContains(band, shifted)) return false;
  }
  return true;
}

function sameRange(left: Range, right: Range): boolean {
  return (
    left.sheet === right.sheet &&
    left.r0 === right.r0 &&
    left.c0 === right.c0 &&
    left.r1 === right.r1 &&
    left.c1 === right.c1
  );
}

function rangeContains(outer: Range, inner: Range): boolean {
  return (
    outer.sheet === inner.sheet &&
    inner.r0 >= outer.r0 &&
    inner.c0 >= outer.c0 &&
    inner.r1 <= outer.r1 &&
    inner.c1 <= outer.c1
  );
}

function commentsFitAfterInsert(
  wb: WorkbookHandle,
  sheet: number,
  wholeRows: boolean,
  insertionAt: number,
  count: number,
): boolean {
  const max = wholeRows ? WHOLE_ROW_END : WHOLE_COLUMN_END;
  const lastMovable = max - count;
  for (const comment of wb.getComments(sheet)) {
    const axis = wholeRows ? comment.row : comment.col;
    if (axis >= insertionAt && axis > lastMovable) return false;
  }
  return true;
}

/**
 * A cut-band move can move cells and the small set of worksheet metadata that
 * `structure.ts` explicitly re-points. The native engine also carries opaque
 * worksheet parts (tables, drawings, pivots, page breaks, and names) whose
 * anchor/reference rewrite is not exposed at this command boundary. Refuse a
 * copied-band cut operation while one of those objects belongs to either sheet;
 * silently leaving its old address is worse than asking the caller to use a
 * narrower operation.
 */
function unsupportedMetadataBlocksBand(
  state: State,
  wb: WorkbookHandle,
  sourceSheet: number,
  targetSheet: number,
): boolean {
  const sheets = new Set([sourceSheet, targetSheet]);

  if (
    state.conditional.rules.some((rule) => sheets.has(rule.range.sheet)) ||
    (state.ui.filterRange !== null && sheets.has(state.ui.filterRange.sheet)) ||
    state.tables.tables.some((table) => sheets.has(table.range.sheet)) ||
    state.charts.charts.some((chart) => sheets.has(chart.source.sheet)) ||
    state.illustrations.illustrations.some((item) => sheets.has(item.sheet)) ||
    state.sparkline.sparklines.size > 0 ||
    state.protection.allowedEditRanges.some((entry) => sheets.has(entry.range.sheet)) ||
    state.sheetViews.views.some((view) => sheets.has(view.sheet)) ||
    state.slicers.slicers.length > 0
  ) {
    return true;
  }

  for (const sheet of sheets) {
    if (wb.getConditionalFormats(sheet).length > 0) return true;
    if ((wb.getSheetAutoFilterXml(sheet) ?? '').trim().length > 0) return true;
    const breaks = wb.getSheetPageBreaks(sheet);
    if (breaks && (breaks.rows.length > 0 || breaks.cols.length > 0)) return true;
  }

  // Names can be workbook-scoped, so there is no safe way to infer the
  // unqualified sheet that an engine formula will bind to after a delete.
  // A name scoped to either involved sheet is equally unsafe to leave stale.
  if (
    [...wb.definedNames()].some((name) => name.localSheetId === -1 || sheets.has(name.localSheetId))
  ) {
    return true;
  }

  if (wb.getTables().some((table) => sheets.has(table.sheetIndex))) return true;
  if (wb.getPivotTables().some((pivot) => sheets.has(pivot.sheetIndex))) return true;

  // Passthrough parts are preserved verbatim, but their anchors are opaque to
  // this command. Reject only known anchored worksheet parts; workbook theme
  // and custom XML passthroughs do not block a band move.
  const anchoredPart =
    /(?:^|\/)(?:charts|drawings|pivot|pivotTables|slicers|tables|externalLinks|connections|queryTables)(?:\/|\.|$)/i;
  if (wb.getPassthroughs().some((part) => anchoredPart.test(part.path))) return true;

  return false;
}

function materializedStateForSheet(
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

function refreshCopiedBandSnapshot(
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
