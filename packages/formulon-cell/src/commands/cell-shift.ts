import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import {
  type ActiveAxisFormatSnapshot,
  captureActiveAxisFormatOverrideForStructure,
  restoreActiveAxisFormatOverrideForStructure,
} from '../engine/axis-format-sync.js';
import {
  captureSyncedFormatKeysForStructure,
  restoreSyncedFormatKeysForStructure,
} from '../engine/synced-format-keys.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import {
  type ActiveWholeSheetFontSnapshot,
  captureActiveWholeSheetFontOverrideForStructure,
  restoreActiveWholeSheetFontOverrideForStructure,
} from '../engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addMergeToMaps } from '../store/merge-maps.js';
import { rangeContainsAddr, rangesIntersect } from '../store/selection-geometry.js';
import {
  type CellFormat,
  type ConditionalRule,
  mutators,
  type RangeValidation,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import { listComments, recordCommentChange } from './comment.js';
import { blockDataTableShift } from './data-table-guard.js';
import { adjustFormulaForCellBandShift } from './formula-refs.js';
import type { History } from './history.js';
import { blockPivotEdit } from './pivot-guard.js';
import { blockedByProtection } from './protection.js';
import {
  recordConditionalRulesChange,
  recordFormatChange,
  recordMergesChangeWithEngine,
} from './slice-history.js';
import { deleteCols, deleteRows, insertCols, insertRows } from './structure.js';

export type InsertCellsDirection = 'down' | 'right';
export type DeleteCellsDirection = 'up' | 'left';

export interface CellRecord {
  addr: Addr;
  value: CellValue;
  formula: string | null;
}

interface ShiftCellRecord extends CellRecord {
  /** Effective XF to carry with a moved value when the source inherited an
   * axis style. Null leaves the destination at its normal/inherited style. */
  styleXf: number | null;
}

interface EngineCommentRecord {
  addr: Addr;
  author: string;
  text: string;
  visible: boolean;
}

export interface ShiftReplayCheckpoint {
  readonly store: State;
  readonly activeAxis: ActiveAxisFormatSnapshot | null;
  readonly activeWholeSheetFont: ActiveWholeSheetFontSnapshot | null;
  readonly syncedFormatKeys: ReadonlySet<string> | null;
  readonly cells: readonly CellRecord[];
  readonly externalFormulas: readonly CellRecord[];
  readonly style: Uint8Array | null;
  readonly comments: readonly EngineCommentRecord[];
  readonly merges: readonly Range[];
}

export interface ShiftHistoryReplayRecovery {
  restore?: (phase: 'undo' | 'redo') => void;
  capture?: () => ShiftReplayCheckpoint;
  before?: ShiftReplayCheckpoint;
  after?: ShiftReplayCheckpoint;
}

const cloneStoreState = (store: SpreadsheetStore): State =>
  structuredClone(store.getState()) as State;

export const captureShiftReplayCheckpoint = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  affected: Range,
): ShiftReplayCheckpoint => {
  const style =
    wb.capabilities.rangeStyleSnapshot === true ? wb.snapshotRangeStyle(affected) : null;
  if (wb.capabilities.rangeStyleSnapshot === true && style === null) {
    throw new Error('cell shift: native range style replay checkpoint failed');
  }
  return {
    store: cloneStoreState(store),
    activeAxis: captureActiveAxisFormatOverrideForStructure(wb, affected.sheet),
    activeWholeSheetFont: captureActiveWholeSheetFontOverrideForStructure(wb, affected.sheet),
    syncedFormatKeys: captureSyncedFormatKeysForStructure(wb, affected.sheet),
    cells: collectSheetCells(wb, affected.sheet),
    externalFormulas: collectExternalFormulaCells(wb, affected.sheet),
    style,
    comments: collectEngineComments(
      wb,
      affected,
      listComments(store.getState(), affected.sheet).map((comment) => comment.addr),
    ),
    merges: wb.capabilities.merges
      ? wb.getMerges(affected.sheet).map((merge) => ({ ...merge }))
      : [],
  };
};

export const restoreShiftReplayCheckpoint = (
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  affected: Range,
  checkpoint: ShiftReplayCheckpoint,
): void => {
  wb.withEngineSyncMuted(() => {
    store.setState(() => structuredClone(checkpoint.store) as State);
    restoreActiveAxisFormatOverrideForStructure(wb, affected.sheet, checkpoint.activeAxis);
    restoreActiveWholeSheetFontOverrideForStructure(
      wb,
      affected.sheet,
      checkpoint.activeWholeSheetFont,
    );
    restoreSyncedFormatKeysForStructure(wb, affected.sheet, checkpoint.syncedFormatKeys);
    wb.withBatchedRecalc(() => {
      restoreSheetCells(wb, affected.sheet, checkpoint.cells);
      restoreFormulaCells(wb, checkpoint.externalFormulas);
    });
    if (checkpoint.style !== null && !wb.restoreRangeStyle(affected.sheet, checkpoint.style)) {
      throw new Error('cell shift: failed to restore transaction style');
    }
    if (wb.capabilities.comments) {
      const currentComments = collectEngineComments(
        wb,
        affected,
        listComments(store.getState(), affected.sheet).map((comment) => comment.addr),
      );
      for (const comment of currentComments) {
        wb.setCommentEntry(comment.addr.sheet, comment.addr.row, comment.addr.col, '', '');
      }
      for (const comment of checkpoint.comments) {
        wb.setCommentEntry(
          comment.addr.sheet,
          comment.addr.row,
          comment.addr.col,
          comment.author,
          comment.text,
          comment.visible,
        );
      }
    }
    if (wb.capabilities.merges) {
      if (!wb.engineClearMerges(affected.sheet)) {
        throw new Error('cell shift: failed to restore merges');
      }
      for (const merge of checkpoint.merges) {
        if (!wb.engineAddMerge(affected.sheet, merge)) {
          throw new Error('cell shift: failed to restore merge');
        }
      }
    }
  });
};

export function insertCells(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: InsertCellsDirection,
): boolean {
  const sheet = range.sheet;
  const lines = wholeLineEdit(store, range);
  if (lines) {
    return lines.axis === 'row'
      ? insertRows(store, wb, history, lines.at, lines.count)
      : insertCols(store, wb, history, lines.at, lines.count);
  }
  if (blockedByProtection(store, sheet, 'insert cells')) return false;
  const affected: Range =
    direction === 'down'
      ? { sheet, r0: range.r0, c0: range.c0, r1: MAX_ROW, c1: range.c1 }
      : { sheet, r0: range.r0, c0: range.c0, r1: range.r1, c1: MAX_COL };
  const delta = direction === 'down' ? range.r1 - range.r0 + 1 : range.c1 - range.c0 + 1;
  if (delta > 0 && shiftWouldOverflow(store, wb, affected, direction, delta)) return false;
  if (blockDataTableShift(wb, affected)) return false;
  // Shifted cells cannot carry a report along.
  if (blockPivotEdit(wb, affected)) return false;
  return shiftCellBand(store, wb, history, affected, direction, delta);
}

export function deleteCells(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  range: Range,
  direction: DeleteCellsDirection,
): boolean {
  const sheet = range.sheet;
  const lines = wholeLineEdit(store, range);
  if (lines) {
    return lines.axis === 'row'
      ? deleteRows(store, wb, history, lines.at, lines.count)
      : deleteCols(store, wb, history, lines.at, lines.count);
  }
  if (blockedByProtection(store, sheet, 'delete cells')) return false;
  const affected: Range =
    direction === 'up'
      ? { sheet, r0: range.r0, c0: range.c0, r1: MAX_ROW, c1: range.c1 }
      : { sheet, r0: range.r0, c0: range.c0, r1: range.r1, c1: MAX_COL };
  const delta = direction === 'up' ? -(range.r1 - range.r0 + 1) : -(range.c1 - range.c0 + 1);
  if (blockDataTableShift(wb, affected)) return false;
  // Shifted cells cannot carry a report along.
  if (blockPivotEdit(wb, affected)) return false;
  return shiftCellBand(store, wb, history, affected, direction === 'up' ? 'down' : 'right', delta);
}

/** A range of whole rows or whole columns on the active sheet. Excel ignores
 *  the shift direction for it and edits the rows or columns themselves. */
function wholeLineEdit(
  store: SpreadsheetStore,
  range: Range,
): { axis: 'row' | 'col'; at: number; count: number } | null {
  if (range.sheet !== store.getState().data.sheetIndex) return null;
  const fullRows = range.c0 === 0 && range.c1 === MAX_COL;
  const fullCols = range.r0 === 0 && range.r1 === MAX_ROW;
  if (fullRows === fullCols) return null;
  return fullRows
    ? { axis: 'row', at: range.r0, count: range.r1 - range.r0 + 1 }
    : { axis: 'col', at: range.c0, count: range.c1 - range.c0 + 1 };
}

function shiftCellBand(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): boolean {
  if (delta === 0) return true;
  const replayRecovery: ShiftHistoryReplayRecovery = {};
  if (history && !history.isReplaying()) {
    replayRecovery.capture = () => captureShiftReplayCheckpoint(wb, store, affected);
  }
  const transaction = history?.begin();
  let transactionClosed = false;
  try {
    if (!shiftCellsWithHistory(wb, history, affected, axis, delta, replayRecovery)) return false;
    const initialStoreCommentAddresses = listComments(store.getState(), affected.sheet)
      .filter((comment) => rangeContainsAddr(affected, comment.addr))
      .map((comment) => comment.addr);
    const initialStoreCommentKeys = new Set(initialStoreCommentAddresses.map(addrKey));
    const engineComments = collectEngineComments(wb, affected, initialStoreCommentAddresses);
    hydrateEngineOnlyComments(store, engineComments, initialStoreCommentKeys);
    const storeCommentAddresses = listComments(store.getState(), affected.sheet)
      .filter((comment) => rangeContainsAddr(affected, comment.addr))
      .map((comment) => comment.addr);
    const commentAddresses = collectShiftCommentAddresses(
      storeCommentAddresses,
      engineComments,
      affected,
      axis,
      delta,
    );
    recordCommentChange(history, store, wb, commentAddresses, () => {
      shiftFormats(store, history, affected, axis, delta);
      shiftEngineComments(wb, affected, axis, delta, engineComments);
    });
    shiftMerges(store, wb, history, affected, axis, delta);
    shiftConditionalRules(store, history, affected, axis, delta);
    wb.recalcAuto();
    if (history && !history.isReplaying()) {
      replayRecovery.restore = (phase) => {
        const checkpoint = phase === 'undo' ? replayRecovery.after : replayRecovery.before;
        if (!checkpoint) throw new Error(`cell shift: ${phase} replay checkpoint was not armed`);
        restoreShiftReplayCheckpoint(wb, store, affected, checkpoint);
      };
      history.push({
        undo: () => {
          replayRecovery.after = replayRecovery.capture?.();
        },
        redo: () => {
          replayRecovery.before = undefined;
          replayRecovery.after = undefined;
        },
      });
    }
  } catch (error) {
    if (transaction !== undefined) {
      transactionClosed = true;
      try {
        history?.abort(transaction);
      } catch (abortError) {
        throw new AggregateError(
          [error, abortError],
          'cell shift: initial transaction rollback failed',
          { cause: error },
        );
      }
    }
    throw error;
  } finally {
    if (transaction !== undefined && !transactionClosed) history?.end(transaction);
  }
  return true;
}

/**
 * Shift a cell band and record the value/formula change as one history child.
 *
 * Cell writes used while moving values can materialize the effective row or
 * column style as a direct XF. The native range checkpoint is deliberately
 * attached to this value child, so both undo and redo restore styles after
 * the value replay has finished. Partial cell shifts keep axis journal keys
 * at their original coordinates; the native checkpoint is the source of
 * truth for the affected rectangle.
 */
export function shiftCellsWithHistory(
  wb: WorkbookHandle,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
  replayRecovery?: ShiftHistoryReplayRecovery,
): boolean {
  const recording = history !== null && !history.isReplaying();
  const beforeCells = recording ? collectSheetCells(wb, affected.sheet) : null;
  const beforeExternalFormulas = recording ? collectExternalFormulaCells(wb, affected.sheet) : null;
  const checkpointEnabled = recording && wb.capabilities.rangeStyleSnapshot === true;
  let beforeStyle: Uint8Array | null = null;

  if (checkpointEnabled) {
    beforeStyle = wb.snapshotRangeStyle(affected);
    // A capable native module must provide a checkpoint. Refusing before the
    // mutation keeps the command atomic when the adapter/native call fails.
    if (beforeStyle === null) return false;
  }

  try {
    shiftCells(wb, affected, axis, delta);
    const afterCells = recording ? collectSheetCells(wb, affected.sheet) : null;
    const afterExternalFormulas = recording
      ? collectExternalFormulaCells(wb, affected.sheet)
      : null;
    let afterStyle: Uint8Array | null = null;
    if (checkpointEnabled) {
      afterStyle = wb.snapshotRangeStyle(affected);
      if (afterStyle === null) {
        throw new Error('cell shift: native range style after-snapshot failed');
      }
    }

    if (history && beforeCells && beforeExternalFormulas && afterCells && afterExternalFormulas) {
      history.push({
        undo: () => {
          if (replayRecovery && replayRecovery.after === undefined) {
            replayRecovery.after = replayRecovery.capture?.();
          }
          replayCellShiftSnapshot(
            wb,
            affected,
            beforeCells,
            beforeExternalFormulas,
            beforeStyle,
            afterCells,
            afterExternalFormulas,
            afterStyle,
            replayRecovery,
            'undo',
          );
        },
        redo: () => {
          if (replayRecovery) replayRecovery.before = replayRecovery.capture?.();
          replayCellShiftSnapshot(
            wb,
            affected,
            afterCells,
            afterExternalFormulas,
            afterStyle,
            beforeCells,
            beforeExternalFormulas,
            beforeStyle,
            replayRecovery,
            'redo',
          );
        },
      });
    }
    return true;
  } catch (error) {
    if (recording && beforeCells && beforeExternalFormulas) {
      const rollbackErrors: unknown[] = [];
      try {
        restoreSheetCells(wb, affected.sheet, beforeCells);
        restoreFormulaCells(wb, beforeExternalFormulas);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
      try {
        restoreCellShiftStyle(wb, affected, beforeStyle);
      } catch (rollbackError) {
        rollbackErrors.push(rollbackError);
      }
      if (rollbackErrors.length > 0) {
        throw new AggregateError(
          [error, ...rollbackErrors],
          'cell shift: value/style rollback failed',
          { cause: error },
        );
      }
    }
    throw error;
  }
}

/** Replay a value/style checkpoint as one atomic history child. Value replay
 * can materialize direct XFs before the range-style restore runs. If that
 * restore rejects, put both the values and the native style back to the state
 * that was visible before this child started; otherwise a failed undo would
 * leave the child half-applied while its containing transaction continues. */
function replayCellShiftSnapshot(
  wb: WorkbookHandle,
  affected: Range,
  targetCells: readonly CellRecord[],
  targetExternalFormulas: readonly CellRecord[],
  targetStyle: Uint8Array | null,
  rollbackCells: readonly CellRecord[],
  rollbackExternalFormulas: readonly CellRecord[],
  rollbackStyle: Uint8Array | null,
  replayRecovery: ShiftHistoryReplayRecovery | undefined,
  phase: 'undo' | 'redo',
): void {
  try {
    restoreSheetCells(wb, affected.sheet, targetCells);
    restoreFormulaCells(wb, targetExternalFormulas);
    restoreCellShiftStyle(wb, affected, targetStyle);
  } catch (error) {
    const rollbackErrors: unknown[] = [];
    try {
      restoreSheetCells(wb, affected.sheet, rollbackCells);
      restoreFormulaCells(wb, rollbackExternalFormulas);
    } catch (rollbackError) {
      rollbackErrors.push(rollbackError);
    }
    try {
      restoreCellShiftStyle(wb, affected, rollbackStyle);
    } catch (rollbackError) {
      rollbackErrors.push(rollbackError);
    }
    try {
      replayRecovery?.restore?.(phase);
    } catch (recoveryError) {
      rollbackErrors.push(recoveryError);
    }
    if (rollbackErrors.length > 0) {
      throw new AggregateError(
        [error, ...rollbackErrors],
        'cell shift: history replay rollback failed',
        { cause: error },
      );
    }
    throw error;
  }
}

function restoreCellShiftStyle(
  wb: WorkbookHandle,
  affected: Range,
  bytes: Uint8Array | null,
): void {
  if (bytes === null) return;
  const restored = wb.withEngineSyncMuted(() => wb.restoreRangeStyle(affected.sheet, bytes));
  if (!restored) throw new Error('cell shift: native range style restore failed');
}

function collectSheetCells(wb: WorkbookHandle, sheet: number): CellRecord[] {
  return Array.from(wb.physicalCells(sheet)).map<CellRecord>((c) => ({
    addr: c.addr,
    value: c.value,
    formula: c.formula,
  }));
}

function collectExternalFormulaCells(wb: WorkbookHandle, editedSheet: number): CellRecord[] {
  const formulas: CellRecord[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    if (sheet === editedSheet) continue;
    for (const cell of wb.physicalCells(sheet)) {
      if (!cell.formula) continue;
      formulas.push({ addr: cell.addr, value: cell.value, formula: cell.formula });
    }
  }
  return formulas;
}

function restoreSheetCells(wb: WorkbookHandle, sheet: number, cells: readonly CellRecord[]): void {
  const existing = Array.from(wb.physicalCells(sheet));
  wb.withBatchedRecalc(() => {
    for (const cell of existing) wb.setBlank(cell.addr);
    for (const cell of cells) writeCell(wb, cell.addr, cell.value, cell.formula);
  });
  wb.recalcAuto();
}

function restoreFormulaCells(wb: WorkbookHandle, cells: readonly CellRecord[]): void {
  wb.withBatchedRecalc(() => {
    for (const cell of cells) {
      if (cell.formula) wb.setFormula(cell.addr, cell.formula);
    }
  });
  wb.recalcAuto();
}

function collectShiftCommentAddresses(
  storeCommentAddresses: readonly Addr[],
  engineComments: readonly EngineCommentRecord[],
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): Addr[] {
  const tracked = new Map<string, Addr>();
  const add = (addr: Addr): void => {
    tracked.set(addrKey(addr), addr);
  };
  for (const addr of storeCommentAddresses) {
    add(addr);
    const shifted = shiftedCommentAddr(addr, axis, delta);
    if (inShiftTarget(shifted, affected, axis)) add(shifted);
  }
  for (const comment of engineComments) {
    add(comment.addr);
    const shifted = shiftedCommentAddr(comment.addr, axis, delta);
    if (inShiftTarget(shifted, affected, axis)) add(shifted);
  }
  return [...tracked.values()];
}

function shiftedCommentAddr(addr: Addr, axis: InsertCellsDirection, delta: number): Addr {
  return axis === 'down' ? { ...addr, row: addr.row + delta } : { ...addr, col: addr.col + delta };
}

function shiftEngineComments(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
  entries: readonly EngineCommentRecord[],
): void {
  if (!wb.capabilities.comments) return;
  if (entries.length === 0) return;

  wb.withBatchedRecalc(() => {
    for (const entry of entries) {
      wb.setCommentEntry(entry.addr.sheet, entry.addr.row, entry.addr.col, '', '');
    }
    for (const entry of entries) {
      const shifted = shiftedCommentAddr(entry.addr, axis, delta);
      if (!inShiftTarget(shifted, affected, axis)) continue;
      wb.setCommentEntry(
        shifted.sheet,
        shifted.row,
        shifted.col,
        entry.author,
        entry.text,
        entry.visible,
      );
    }
  });
}

function collectEngineComments(
  wb: WorkbookHandle,
  affected: Range,
  storeCommentAddresses: readonly Addr[],
): EngineCommentRecord[] {
  const entries = wb.capabilities.commentsEnumerable
    ? wb.getComments(affected.sheet).map((comment) => ({
        addr: { sheet: affected.sheet, row: comment.row, col: comment.col },
        author: comment.author,
        text: comment.text,
        visible: comment.visible === true,
      }))
    : storeCommentAddresses.flatMap((addr) => {
        const comment = wb.getComment(addr.sheet, addr.row, addr.col);
        return comment
          ? [
              {
                addr,
                author: comment.author,
                text: comment.text,
                visible: comment.visible === true,
              },
            ]
          : [];
      });
  const unique = new Map<string, EngineCommentRecord>();
  for (const entry of entries) {
    if (rangeContainsAddr(affected, entry.addr)) unique.set(addrKey(entry.addr), entry);
  }
  return [...unique.values()];
}

function hydrateEngineOnlyComments(
  store: SpreadsheetStore,
  entries: readonly EngineCommentRecord[],
  storeCommentKeys: ReadonlySet<string>,
): void {
  for (const entry of entries) {
    if (storeCommentKeys.has(addrKey(entry.addr))) continue;
    mutators.setCellFormat(store, entry.addr, {
      comment: entry.text,
      commentAuthor: entry.author || undefined,
      commentVisible: entry.visible || undefined,
    });
  }
}

/** Move the cells inside `affected` by `delta` along `axis` and rewrite every
 *  formula on every sheet that references the moved band. */
export function shiftCells(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  wb.withBatchedRecalc(() => writeShiftedCells(wb, affected, axis, delta));
}

function writeShiftedCells(
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  const allSheets = Array.from({ length: wb.sheetCount }, (_, sheet) =>
    Array.from(wb.physicalCells(sheet)).map<ShiftCellRecord>((c) => ({
      addr: c.addr,
      value: c.value,
      formula: c.formula,
      styleXf: styleXfForMovedCell(wb, c.addr),
    })),
  );
  const all = allSheets[affected.sheet] ?? [];
  const sheetNames = Array.from({ length: wb.sheetCount }, (_, sheet) => wb.sheetName(sheet));
  const moving = all.filter((cell) => rangeContainsAddr(affected, cell.addr));
  for (const cell of moving) {
    wb.setBlank(cell.addr);
    // setBlank clears contents only. A partial shift must also remove the
    // source's direct XF, including an explicit blank XF0 or an imported XF
    // whose value has moved elsewhere. Older native bundles cannot express
    // that distinction, so keep their historical behavior behind the
    // capability guard.
    if (
      wb.capabilities.cellFormattingClear === true &&
      !wb.clearCellXfIndex(cell.addr.sheet, cell.addr.row, cell.addr.col)
    ) {
      throw new Error('cell shift: failed to clear moved source style');
    }
  }

  const sorted =
    axis === 'down'
      ? moving.sort((a, b) => (delta > 0 ? b.addr.row - a.addr.row : a.addr.row - b.addr.row))
      : moving.sort((a, b) => (delta > 0 ? b.addr.col - a.addr.col : a.addr.col - b.addr.col));

  for (const cell of sorted) {
    const next: Addr =
      axis === 'down'
        ? { ...cell.addr, row: cell.addr.row + delta }
        : { ...cell.addr, col: cell.addr.col + delta };
    if (!inShiftTarget(next, affected, axis)) continue;
    const formula = cell.formula
      ? adjustFormulaForCellBandShift(cell.formula, affected, axis, delta, {
          editedSheet: affected.sheet,
          formulaSheet: affected.sheet,
          sheetNames,
        })
      : null;
    writeCell(wb, next, cell.value, formula);
    if (cell.styleXf !== null && !wb.setCellXfIndex(next.sheet, next.row, next.col, cell.styleXf)) {
      throw new Error('cell shift: failed to preserve moved cell style');
    }
  }

  for (let sheet = 0; sheet < allSheets.length; sheet += 1) {
    const formulaSheet = allSheets[sheet] ?? [];
    for (const cell of formulaSheet) {
      if (!cell.formula || (sheet === affected.sheet && rangeContainsAddr(affected, cell.addr)))
        continue;
      const nextFormula = adjustFormulaForCellBandShift(cell.formula, affected, axis, delta, {
        editedSheet: affected.sheet,
        formulaSheet: sheet,
        sheetNames,
      });
      if (nextFormula !== cell.formula) wb.setFormula(cell.addr, nextFormula);
    }
  }
}

/** Preserve the visible style of a value while moving it between rows or
 * columns. A partial cell shift does not move the row/column axis records, so
 * writing only the value would otherwise resolve the destination against a
 * different inherited XF. */
function styleXfForMovedCell(wb: WorkbookHandle, addr: Addr): number | null {
  const effective = wb.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  if (effective?.source === 0) return effective.xfIndex;
  if (effective && effective.xfIndex !== 0) return effective.xfIndex;
  return null;
}

export function shiftFormats(
  store: SpreadsheetStore,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  recordFormatChange(history, store, () => {
    store.setState((s) => {
      const shifted = shiftFormatMap(s.format.formats, affected, axis, delta);
      const formats =
        delta > 0
          ? inheritInsertedValidations(s.format.formats, shifted, affected, axis, delta)
          : shifted;
      const rangeValidations = shiftRangeValidationsForBand(
        s.format.rangeValidations,
        affected,
        axis,
        delta,
      );
      return { ...s, format: { ...s.format, formats, rangeValidations } };
    });
  });
}

export function shiftMerges(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  recordMergesChangeWithEngine(history, store, wb, affected.sheet, () => {
    store.setState((s) => {
      const byAnchor = new Map<string, Range>();
      const byCell = new Map<string, string>();
      for (const merge of s.merges.byAnchor.values()) {
        // A merge the band cuts through is unmerged; one inside it moves whole.
        const inBand =
          merge.sheet === affected.sheet && mergeIntersectsShiftBand(merge, affected, axis);
        if (inBand && !mergeFitsBand(merge, affected, axis)) continue;
        const shifted = inBand ? shiftRange(merge, axis, delta) : merge;
        if (shifted.r0 < 0 || shifted.c0 < 0 || shifted.r1 > MAX_ROW || shifted.c1 > MAX_COL) {
          continue;
        }
        addMergeToMaps(byAnchor, byCell, shifted);
      }
      return { ...s, merges: { byAnchor, byCell } };
    });
  });
}

/** Whether every merge touching the shift band moves whole; false when a
 *  merge straddles the band edge and would be split. */
export function canShiftMerges(state: State, affected: Range, axis: InsertCellsDirection): boolean {
  for (const merge of state.merges.byAnchor.values()) {
    if (merge.sheet !== affected.sheet) continue;
    if (!rangesIntersect(merge, affected)) continue;
    if (!mergeIntersectsShiftBand(merge, affected, axis)) continue;
    if (!mergeFitsBand(merge, affected, axis)) return false;
  }
  return true;
}

function mergeFitsBand(merge: Range, affected: Range, axis: InsertCellsDirection): boolean {
  return axis === 'down'
    ? merge.c0 >= affected.c0 && merge.c1 <= affected.c1 && merge.r0 >= affected.r0
    : merge.r0 >= affected.r0 && merge.r1 <= affected.r1 && merge.c0 >= affected.c0;
}

function shiftWouldOverflow(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): boolean {
  const overflows = (addr: Addr): boolean => {
    if (!rangeContainsAddr(affected, addr)) return false;
    return axis === 'down' ? addr.row + delta > MAX_ROW : addr.col + delta > MAX_COL;
  };
  for (const cell of wb.physicalCells(affected.sheet)) {
    if (overflows(cell.addr)) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — content would leave the worksheet');
      return true;
    }
  }
  for (const key of store.getState().format.formats.keys()) {
    const addr = parseAddrKey(key);
    if (!addr) continue;
    if (overflows(addr)) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — format would leave the worksheet');
      return true;
    }
  }
  if (wb.capabilities.commentsEnumerable) {
    for (const comment of wb.getComments(affected.sheet)) {
      if (overflows({ sheet: affected.sheet, row: comment.row, col: comment.col })) {
        // eslint-disable-next-line no-console
        console.warn('formulon-cell: cell shift blocked — comment would leave the worksheet');
        return true;
      }
    }
  }
  for (const merge of store.getState().merges.byAnchor.values()) {
    if (
      merge.sheet !== affected.sheet ||
      !mergeIntersectsShiftBand(merge, affected, axis) ||
      !mergeFitsBand(merge, affected, axis)
    ) {
      continue;
    }
    const shifted = shiftRange(merge, axis, delta);
    if (shifted.r1 > MAX_ROW || shifted.c1 > MAX_COL) {
      // eslint-disable-next-line no-console
      console.warn('formulon-cell: cell shift blocked — merge would leave the worksheet');
      return true;
    }
  }
  return false;
}

const transposeRange = (r: Range): Range => ({ ...r, r0: r.c0, c0: r.r0, r1: r.c1, c1: r.r1 });

/** Areas a range covers once the cells of `affected` move by `delta`. Cells
 *  outside the band stay; band cells move, and deleted ones vanish. Inserted
 *  cells join a range that has the cell before the band (`edge`) or both
 *  neighbours of the insertion point (`inside`). */
function rangePiecesAfterBandShift(
  range: Range,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
  inherit: 'inside' | 'edge',
): Range[] {
  if (range.sheet !== affected.sheet) return [range];
  if (axis === 'right') {
    return rangePiecesAfterBandShift(
      transposeRange(range),
      transposeRange(affected),
      'down',
      delta,
      inherit,
    ).map(transposeRange);
  }
  const m0 = Math.max(range.c0, affected.c0);
  const m1 = Math.min(range.c1, affected.c1);
  if (m0 > m1 || range.r1 < affected.r0) return [range];
  const pieces: Range[] = [];
  if (range.c0 < m0) pieces.push({ ...range, c1: m0 - 1 });
  if (range.c1 > m1) pieces.push({ ...range, c0: m1 + 1 });
  if (range.r0 < affected.r0) {
    pieces.push({ ...range, c0: m0, c1: m1, r1: Math.min(range.r1, affected.r0 - 1) });
  }
  const moved = (from: number, to: number): void => {
    const r0 = from + delta;
    const r1 = Math.min(to + delta, MAX_ROW);
    if (r0 <= r1) pieces.push({ ...range, c0: m0, c1: m1, r0, r1 });
  };
  if (delta > 0) {
    moved(Math.max(range.r0, affected.r0), range.r1);
    const above = affected.r0 - 1;
    if (range.r0 <= above && above <= range.r1 && (inherit === 'edge' || range.r1 >= affected.r0)) {
      pieces.push({
        ...range,
        c0: m0,
        c1: m1,
        r0: affected.r0,
        r1: Math.min(affected.r0 + delta - 1, MAX_ROW),
      });
    }
  } else {
    const from = Math.max(range.r0, affected.r0 - delta);
    if (from <= range.r1) moved(from, range.r1);
    else if (inherit === 'inside') {
      // A deleted block that reaches the range's trailing edge leaves the rule whole.
      pieces.push({ ...range, c0: m0, c1: m1, r0: Math.max(range.r0, affected.r0) });
    }
  }
  return coalesceRanges(pieces);
}

/** Merge pieces that share a full edge into single rectangles. */
function coalesceRanges(pieces: Range[]): Range[] {
  const out = pieces.map((p) => ({ ...p }));
  let merged = true;
  while (merged) {
    merged = false;
    for (let i = 0; i < out.length && !merged; i += 1) {
      for (let j = i + 1; j < out.length && !merged; j += 1) {
        const a = out[i] as Range;
        const b = out[j] as Range;
        const sameCols = a.c0 === b.c0 && a.c1 === b.c1;
        const sameRows = a.r0 === b.r0 && a.r1 === b.r1;
        if (sameCols && a.r0 <= b.r1 + 1 && b.r0 <= a.r1 + 1) {
          out[i] = { ...a, r0: Math.min(a.r0, b.r0), r1: Math.max(a.r1, b.r1) };
        } else if (sameRows && a.c0 <= b.c1 + 1 && b.c0 <= a.c1 + 1) {
          out[i] = { ...a, c0: Math.min(a.c0, b.c0), c1: Math.max(a.c1, b.c1) };
        } else continue;
        out.splice(j, 1);
        merged = true;
      }
    }
  }
  return out.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
}

const sameRange = (a: Range, b: Range): boolean =>
  a.sheet === b.sheet && a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;

/** Re-point conditional-format ranges after a cell shift; a rule whose cells
 *  end up in several areas becomes one rule per area. */
function shiftConditionalRules(
  store: SpreadsheetStore,
  history: History | null,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): void {
  recordConditionalRulesChange(history, store, () => {
    store.setState((s) => {
      let changed = false;
      const rules = s.conditional.rules.flatMap<ConditionalRule>((rule) => {
        if (rule.range.sheet !== affected.sheet) return [rule];
        const pieces = rangePiecesAfterBandShift(rule.range, affected, axis, delta, 'inside');
        const [only] = pieces;
        if (pieces.length === 1 && only && sameRange(only, rule.range)) return [rule];
        changed = true;
        return pieces.map((range) => ({ ...rule, range }));
      });
      return changed ? { ...s, conditional: { rules } } : s;
    });
  });
}

/** Range-stored validations follow the cells of the band; inserted cells take
 *  the rule of the cell above (or to the left). */
function shiftRangeValidationsForBand(
  rules: RangeValidation[] | undefined,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): RangeValidation[] | undefined {
  if (!rules?.some((rule) => rule.ranges.some((range) => range.sheet === affected.sheet))) {
    return rules;
  }
  return rules.flatMap((rule) => {
    const ranges = rule.ranges.flatMap((range) =>
      rangePiecesAfterBandShift(range, affected, axis, delta, 'edge'),
    );
    return ranges.length > 0 ? [{ ...rule, ranges }] : [];
  });
}

/** Inserted cells take the validation of the cell above (or to the left). */
function inheritInsertedValidations(
  before: Map<string, CellFormat>,
  shifted: Map<string, CellFormat>,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): Map<string, CellFormat> {
  let out = shifted;
  const down = axis === 'down';
  const alongStart = down ? affected.r0 : affected.c0;
  const acrossLo = down ? affected.c0 : affected.r0;
  const acrossHi = down ? affected.c1 : affected.r1;
  for (const [key, fmt] of before) {
    if (!fmt.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== affected.sheet) continue;
    const along = down ? addr.row : addr.col;
    const across = down ? addr.col : addr.row;
    if (along !== alongStart - 1 || across < acrossLo || across > acrossHi) continue;
    if (out === shifted) out = new Map(shifted);
    for (
      let at = alongStart;
      at < alongStart + delta && at <= (down ? MAX_ROW : MAX_COL);
      at += 1
    ) {
      const target = addrKey(down ? { ...addr, row: at } : { ...addr, col: at });
      out.set(target, { ...out.get(target), validation: fmt.validation });
    }
  }
  return out;
}

function shiftFormatMap(
  formats: Map<string, CellFormat>,
  affected: Range,
  axis: InsertCellsDirection,
  delta: number,
): Map<string, CellFormat> {
  const next = new Map<string, CellFormat>();
  for (const [key, fmt] of formats) {
    const addr = parseAddrKey(key);
    if (!addr) {
      next.set(key, fmt);
      continue;
    }
    if (!rangeContainsAddr(affected, addr)) {
      next.set(key, fmt);
      continue;
    }
    const shifted =
      axis === 'down' ? { ...addr, row: addr.row + delta } : { ...addr, col: addr.col + delta };
    if (inShiftTarget(shifted, affected, axis)) next.set(addrKey(shifted), fmt);
  }
  return next;
}

function inShiftTarget(addr: Addr, affected: Range, axis: InsertCellsDirection): boolean {
  if (addr.sheet !== affected.sheet || addr.row < 0 || addr.col < 0) return false;
  if (addr.row > MAX_ROW || addr.col > MAX_COL) return false;
  return axis === 'down'
    ? addr.row >= affected.r0 && addr.col >= affected.c0 && addr.col <= affected.c1
    : addr.col >= affected.c0 && addr.row >= affected.r0 && addr.row <= affected.r1;
}

function mergeIntersectsShiftBand(
  merge: Range,
  affected: Range,
  axis: InsertCellsDirection,
): boolean {
  return axis === 'down'
    ? merge.r1 >= affected.r0 && merge.c1 >= affected.c0 && merge.c0 <= affected.c1
    : merge.c1 >= affected.c0 && merge.r1 >= affected.r0 && merge.r0 <= affected.r1;
}

function shiftRange(range: Range, axis: InsertCellsDirection, delta: number): Range {
  return axis === 'down'
    ? { ...range, r0: range.r0 + delta, r1: range.r1 + delta }
    : { ...range, c0: range.c0 + delta, c1: range.c1 + delta };
}

// Reference shifting for cell-band inserts/deletes is delegated to the shared
// `adjustFormulaForCellBandShift` (see commands/formula-refs.ts), which handles
// sheet qualifiers, function names, ranges, and grid bounds uniformly.
