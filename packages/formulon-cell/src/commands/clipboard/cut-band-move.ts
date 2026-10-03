import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import { writeCell } from '../../engine/value.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { addMergeToMaps, removeIntersectingMerges } from '../../store/merge-maps.js';
import { rangeContainsRange, rangesIntersect } from '../../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../../store/store.js';
import type { CellRecord } from '../cell-shift.js';
import { collectAllFormulas, type FormulaRecord } from '../formula-records.js';
import { type AxisBandMoveContext, adjustFormulaForAxisBandMove } from '../formula-refs.js';
import type { History } from '../history.js';
import { isSheetProtected } from '../protection.js';
import { recordMergesChangeWithEngine } from '../slice-history.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../structure.js';
import {
  cloneCellFormat,
  commentsFitAfterInsert,
  makeBand,
  materializedStateForSheet,
  refreshCopiedBandSnapshot,
} from './copied-band.js';
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
import type { InsertCopiedCellsResult } from './insert-copied-cells.js';
import { pasteSpecial } from './paste-special.js';
import type { ClipboardSnapshot } from './snapshot.js';

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

export function finalStartForCut(sourceStart: number, count: number, insertionAt: number): number {
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
      const addr = parseAddrKey(key);
      if (!addr) continue;
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
    if (rangesIntersect(merge, source) && !rangeContainsRange(source, merge)) return false;
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
  const max = wholeRows ? MAX_ROW : MAX_COL;
  const finalStart = finalStartForCut(sourceStart, count, insertionAt);
  if (finalStart < 0 || finalStart + count > max) return false;
  const mappedIndex = (index: number): number =>
    mapCutAxisIndex(index, sourceStart, count, insertionAt);

  for (const cell of wb.cells(targetSheet)) {
    const index = axis === 'row' ? cell.addr.row : cell.addr.col;
    if (mappedIndex(index) > max) return false;
  }
  for (const key of state.format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== targetSheet) continue;
    const index = wholeRows ? addr.row : addr.col;
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
    const addr = parseAddrKey(key);
    if (!addr) continue;
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

export function insertCutCopiedBand(
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
  const originalFormulas = snapshot.range.sheet === targetSheet ? collectAllFormulas(wb) : [];
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

/**
 * A cut-band move can move cells and the small set of worksheet metadata that
 * `structure.ts` explicitly re-points. The native engine also carries opaque
 * worksheet parts (tables, drawings, pivots, page breaks, and names) whose
 * anchor/reference rewrite is not exposed at this command boundary. Refuse a
 * copied-band cut operation while one of those objects belongs to either sheet;
 * silently leaving its old address is worse than asking the caller to use a
 * narrower operation.
 */
export function unsupportedMetadataBlocksBand(
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
