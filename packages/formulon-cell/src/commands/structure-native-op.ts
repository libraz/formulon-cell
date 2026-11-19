import { AnchorEditAs, DrawingObjectKind } from '@libraz/formulon';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import { restoreNativeAxisSnapshot } from '../engine/axis-format-sync.js';
import { type AxisStructureShift, mapAxisAddress } from '../engine/axis-structure-remap.js';
import { hydrateChartsFromEngine, hydrateRestoredCharts } from '../engine/charts-sync.js';
import {
  type FormatHistoryTrackingSnapshot,
  formatHistoryBridgeFor,
  type NativeFormatSnapshot,
} from '../engine/format-history-bridge.js';
import { hydrateImagesFromEngine, hydrateRestoredImages } from '../engine/images-sync.js';
import { hydrateProtectedRangesFromEngine } from '../engine/protection-sync.js';
import { captureShapeAnchors } from '../engine/shapes-sync.js';
import {
  captureSlicerAnchors,
  captureSlicerRestore,
  hydrateSlicersFromEngine,
} from '../engine/slicer-sync.js';
import { captureSparklineGroups, hydrateSparklinesFromEngine } from '../engine/sparkline-sync.js';
import type { Range } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { EngineDrawingObject } from '../engine/workbook-handle-annotations.js';
import { sheetDrawingStack, writeDrawingOrder } from '../store/drawing-order.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { CellRecord } from './cell-shift.js';
import { collectAllFormulas, type FormulaRecord } from './formula-records.js';
import type { History } from './history.js';
import {
  captureStructureReplayCheckpoint,
  type StructureReplayCheckpointRef,
  withEngineSyncMuted,
} from './structure-checkpoint.js';

const structureStyleRange = (sheet: number, axis: 'row' | 'col', split: number): Range => ({
  sheet,
  r0: axis === 'row' ? split : 0,
  c0: axis === 'col' ? split : 0,
  r1: MAX_ROW,
  c1: MAX_COL,
});

const axisSnapshotFor = (
  wb: WorkbookHandle,
  sheet: number,
): ReturnType<WorkbookHandle['getSheetFontAxisSnapshot']> => {
  const candidate = wb as WorkbookHandle & {
    getSheetFontAxisSnapshot?: WorkbookHandle['getSheetFontAxisSnapshot'];
  };
  return typeof candidate.getSheetFontAxisSnapshot === 'function'
    ? candidate.getSheetFontAxisSnapshot(sheet)
    : null;
};

/**
 * Apply a row/col shift using the engine's native `insertRows/deleteRows/
 * insertCols/deleteCols` ops. Pushes one history entry that inverts via the
 * opposite engine op (with a captured cell-band restore for delete). Cells
 * outside the shifted band are *not* touched in the store cache here; the
 * caller is expected to refresh via `replaceCells` after this returns.
 */
/** Whether deleting `count` lines at `split` removes the drawing object: the
 * engine drops a move-and-size object whose both corners lie in the band. */
const drawingRemovedByDelete = (
  object: EngineDrawingObject,
  axis: 'row' | 'col',
  split: number,
  count: number,
): boolean => {
  const inBand = (line: number): boolean => line >= split && line < split + count;
  return (
    object.anchor.editAs === AnchorEditAs.TwoCell &&
    inBand(axis === 'row' ? object.anchor.fromRow : object.anchor.fromCol) &&
    inBand(axis === 'row' ? object.anchor.toRow : object.anchor.toCol)
  );
};

/** A row delete the engine runs over more rows than it removes: deleting every data row of a
 *  table keeps the first one, blank, so the engine deletes `count` rows from `split` and the
 *  sheet loses only the ones below it. */
interface NativeRowDelete {
  readonly split: number;
  readonly count: number;
}

export function applyAxisShiftViaEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  history: History | null,
  sheet: number,
  axis: 'row' | 'col',
  split: number,
  delta: number,
  restoreAfterFailedReplay?: (phase: 'redo' | 'undo') => void,
  replayCheckpoint?: StructureReplayCheckpointRef,
  nativeDelete?: NativeRowDelete,
): boolean {
  if (delta === 0 && !nativeDelete) return true;
  const positive = delta > 0;
  const count = Math.abs(delta);
  // The rows the engine is asked to delete; the shift itself is `split`/`delta`.
  const deleteSplit = nativeDelete?.split ?? split;
  const deleteCount = nativeDelete?.count ?? count;
  const record = history !== null && !history.isReplaying();
  const shift: AxisStructureShift = { axis, split, delta };
  const bridge = record ? formatHistoryBridgeFor(store) : undefined;
  const styleRange = structureStyleRange(sheet, axis, deleteSplit);
  const rangeStyleCapable = record && wb.capabilities.rangeStyleSnapshot === true;
  const beforeRangeStyle = rangeStyleCapable ? wb.snapshotRangeStyle(styleRange) : null;
  // A capable native module must produce the opaque checkpoint before any
  // structural mutation. Falling back to the manual sparse checkpoint here
  // would silently lose blank-cell XFs and axis metadata on Undo.
  if (rangeStyleCapable && beforeRangeStyle === null) return false;
  const useRangeStyleCheckpoint = rangeStyleCapable;

  // A structure edit moves sparse direct XFs even when the store has no
  // projected CellFormat for them. Capture the source keys before native
  // mutation and their mapped destinations after it. Deleted-band keys stay in
  // the before checkpoint so Undo can recreate them after the native inverse.
  const sourceKeys = new Set<string>();
  if (record) {
    for (const key of store.getState().format.formats.keys()) {
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== sheet) continue;
      const index = axis === 'row' ? addr.row : addr.col;
      if (index >= deleteSplit) sourceKeys.add(key);
    }
    for (const entry of store.getState().format.sheetAxisOverrides?.get(sheet)?.directXfs ?? []) {
      sourceKeys.add(entry.key);
    }
  }
  const beforeNative: NativeFormatSnapshot =
    record && bridge && !useRangeStyleCheckpoint ? bridge.capture(sourceKeys) : new Map();
  const beforeTracking: FormatHistoryTrackingSnapshot =
    record && bridge ? bridge.captureTracking(sourceKeys) : new Map();
  const beforeAxes = record && !useRangeStyleCheckpoint ? axisSnapshotFor(wb, sheet) : null;
  let afterNative: NativeFormatSnapshot = new Map();
  let afterTracking: FormatHistoryTrackingSnapshot = new Map();
  let afterAxes = beforeAxes;
  let afterRangeStyle: Uint8Array | null = null;

  // For delete (delta < 0), capture cells in the band so undo can rewrite
  // them. Insert is fully invertible by the opposite engine op alone.
  const captured: CellRecord[] = [];
  const originalFormulas: FormulaRecord[] = [];
  if (!positive && record) {
    for (const c of wb.cells(sheet)) {
      const k = axis === 'row' ? c.addr.row : c.addr.col;
      if (k >= deleteSplit && k < deleteSplit + deleteCount) {
        captured.push({ addr: c.addr, value: c.value, formula: c.formula });
      }
    }
    originalFormulas.push(...collectAllFormulas(wb));
  }
  // A delete drops or trims the protected ranges inside the band; undo writes them back.
  const protectedRangesBefore =
    !positive && record && wb.capabilities.protectedRanges ? wb.getProtectedRanges(sheet) : [];
  // The engine deletes the pictures, charts and shapes lying inside the band; undo
  // puts them back from these captures, under their own ids and in their slots.
  const deletedDrawings: Uint8Array[] = [];
  // Table slicers have no snapshot; they, and any drawing without one, are inserted again.
  const deletedObjects: (() => boolean)[] = [];
  // The store id of each removed drawing by object id, and the stacking order to put back.
  const deletedIds = new Map<number, string>();
  const orderBefore = store.getState().drawings.order;
  if (!positive && record && wb.capabilities.drawings) {
    const stack = sheetDrawingStack(store.getState(), sheet);
    for (const object of wb.listDrawingObjects(sheet)) {
      if (!drawingRemovedByDelete(object, axis, split, count)) continue;
      const entry = stack.find(
        (candidate) =>
          candidate.objectId === object.objectId &&
          candidate.chart === (object.kind === DrawingObjectKind.Chart),
      );
      if (entry) deletedIds.set(object.objectId, entry.id);
      const snapshot = wb.snapshotDrawing(sheet, object.objectId);
      if (snapshot !== null) {
        deletedDrawings.push(snapshot);
        continue;
      }
      const restore = captureSlicerRestore(wb, sheet, object.objectId);
      if (restore === null) return false;
      deletedObjects.push(restore);
    }
  }

  // A delete collapses the shapes and slicers it does not remove.
  const restoreAnchors =
    !positive && record ? [captureShapeAnchors(wb, sheet), captureSlicerAnchors(wb, sheet)] : [];
  // The engine drops the sparklines whose cell a delete removes.
  const restoreSparklines = !positive && record ? captureSparklineGroups(wb, store) : null;

  const canRestoreMerges = record && wb.capabilities.merges;
  const beforeMerges = canRestoreMerges ? wb.getMerges(sheet) : [];
  let afterMerges: Range[] = [];
  let initialApply = true;

  const runNativeOp = (insert: boolean): boolean => {
    let ok = false;
    if (insert) {
      // Undoing a delete that removed no row (the table's only data row, blanked) inserts none.
      // Undoing one that kept a table's first data row inserts at that row, inside the table, so
      // the engine grows the table back; the captured band then refills every row.
      const at = nativeDelete ? deleteSplit : split;
      ok =
        count === 0 ||
        (axis === 'row'
          ? wb.engineInsertRows(sheet, at, count)
          : wb.engineInsertCols(sheet, at, count));
    } else {
      ok =
        axis === 'row'
          ? wb.engineDeleteRows(sheet, deleteSplit, deleteCount)
          : wb.engineDeleteCols(sheet, deleteSplit, deleteCount);
    }
    if (!ok) return false;
    wb.recalcAuto();
    hydrateChartsFromEngine(wb, store, sheet);
    hydrateImagesFromEngine(wb, store, sheet);
    hydrateSparklinesFromEngine(wb, store);
    hydrateSlicersFromEngine(wb, store);
    hydrateProtectedRangesFromEngine(wb, store, sheet);
    return true;
  };

  const restoreMerges = (snapshot: readonly Range[]): void => {
    if (!canRestoreMerges) return;
    if (!wb.engineClearMerges(sheet)) throw new Error('structure: failed to restore merges');
    for (const merge of snapshot) {
      if (!wb.engineAddMerge(sheet, merge)) throw new Error('structure: failed to restore merge');
    }
  };

  const afterKeys = (): Set<string> => {
    const keys = new Set<string>();
    for (const key of sourceKeys) {
      const addr = parseAddrKey(key);
      if (!addr) continue;
      const mapped = mapAxisAddress(addr, shift);
      if (mapped) keys.add(addrKey(mapped));
    }
    return keys;
  };

  const restoreCheckpoint = (
    axes: ReturnType<WorkbookHandle['getSheetFontAxisSnapshot']>,
    native: NativeFormatSnapshot,
    tracking: FormatHistoryTrackingSnapshot,
    rangeStyle: Uint8Array | null = null,
  ): void => {
    const errors: unknown[] = [];
    const currentAxes = axes !== null ? axisSnapshotFor(wb, sheet) : null;
    let currentRangeStyle: Uint8Array | null = null;
    if (rangeStyle !== null) {
      if (wb.capabilities.rangeStyleSnapshot !== true) {
        errors.push(new Error('structure: native range style snapshot capability unavailable'));
      } else {
        try {
          currentRangeStyle = wb.snapshotRangeStyle(styleRange);
          if (currentRangeStyle === null)
            throw new Error('structure: native range style rollback snapshot failed');
        } catch (error) {
          errors.push(error);
        }
      }
    }
    let currentNative: NativeFormatSnapshot | null = null;
    let currentTracking: FormatHistoryTrackingSnapshot | null = null;
    if (bridge) {
      try {
        currentNative = bridge.capture(native.keys());
        currentTracking = bridge.captureTracking(native.keys());
      } catch (error) {
        // Continue with the requested restore. If it fails too, the capture
        // error is retained alongside the setter error for the caller.
        errors.push(error);
      }
    }
    let axisFailed = false;
    let directFailed = false;
    let rangeFailed = false;
    const restore = (): void => {
      // A null range checkpoint means this is the legacy/manual path. Keep
      // the axis/direct-XF restore active for older native bundles instead of
      // treating the absence of a range snapshot as a successful restore.
      let rangeRestored = false;
      if (rangeStyle !== null) {
        try {
          if (
            wb.capabilities.rangeStyleSnapshot !== true ||
            !wb.restoreRangeStyle(sheet, rangeStyle)
          )
            throw new Error('structure: failed to restore native range styles');
          rangeRestored = true;
        } catch (error) {
          rangeFailed = true;
          errors.push(error);
        }
      }
      if (!rangeRestored) {
        // Axis and direct-XF recovery are independent. A failed axis setter
        // must not short-circuit a direct checkpoint restore (and vice versa).
        if (axes !== null) {
          try {
            if (!restoreNativeAxisSnapshot(wb, sheet, axes))
              throw new Error('structure: failed to restore native axis styles');
          } catch (error) {
            axisFailed = true;
            errors.push(error);
          }
        }
        if (bridge) {
          try {
            bridge.restore(native);
          } catch (error) {
            directFailed = true;
            errors.push(error);
          }
          if (!axisFailed && !directFailed) {
            try {
              bridge.restoreTracking(tracking);
            } catch (error) {
              errors.push(error);
            }
          }
        }
      } else if (bridge) {
        try {
          bridge.restoreTracking(tracking);
        } catch (error) {
          errors.push(error);
        }
      }
    };
    if (bridge) bridge.withSyncMuted(restore);
    else withEngineSyncMuted(wb, restore);
    if (axisFailed || directFailed) {
      const rollback = (): void => {
        if ((rangeFailed || errors.length > 0) && currentRangeStyle !== null) {
          try {
            if (!wb.restoreRangeStyle(sheet, currentRangeStyle))
              throw new Error('structure: failed to roll back native range styles');
          } catch (error) {
            errors.push(error);
          }
        }
        if (axisFailed && currentAxes !== null) {
          try {
            if (!restoreNativeAxisSnapshot(wb, sheet, currentAxes))
              throw new Error('structure: failed to roll back native axis styles');
          } catch (error) {
            errors.push(error);
          }
        }
        if (directFailed && bridge && currentNative !== null) {
          try {
            bridge.restore(currentNative);
            if (currentTracking !== null) bridge.restoreTracking(currentTracking);
          } catch (error) {
            errors.push(error);
          }
        }
      };
      if (bridge) bridge.withSyncMuted(rollback);
      else withEngineSyncMuted(wb, rollback);
    } else if (errors.length > 0 && currentRangeStyle !== null) {
      const rollbackRange = (): void => {
        try {
          if (!wb.restoreRangeStyle(sheet, currentRangeStyle as Uint8Array))
            throw new Error('structure: failed to roll back native range styles');
        } catch (error) {
          errors.push(error);
        }
        if (bridge && currentTracking !== null) {
          try {
            bridge.restoreTracking(currentTracking);
          } catch (error) {
            errors.push(error);
          }
        }
      };
      if (bridge) bridge.withSyncMuted(rollbackRange);
      else withEngineSyncMuted(wb, rollbackRange);
    }
    if (errors.length === 1) throw errors[0];
    if (errors.length > 1) {
      throw new AggregateError(errors, 'structure: native format checkpoint restore failed', {
        cause: errors[0],
      });
    }
  };

  const restoreDeletedBand = (): void => {
    if (positive) return;
    wb.withBatchedRecalc(() => {
      for (const c of captured) writeCell(wb, c.addr, c.value, c.formula);
      for (const entry of originalFormulas) wb.setFormula(entry.addr, entry.formula);
    });
    restoreSparklines?.();
    if (protectedRangesBefore.length > 0) {
      for (const record of protectedRangesBefore) {
        const index = wb.getProtectedRanges(sheet).findIndex((r) => r.name === record.name);
        if (index >= 0) wb.updateProtectedRange(sheet, index, record);
        else wb.addProtectedRange(sheet, record);
      }
      hydrateProtectedRangesFromEngine(wb, store, sheet);
    }
    // Survivors first, so a restored object cannot take a survivor's captured anchor.
    for (const restore of restoreAnchors) restore();
    for (const snapshot of deletedDrawings) {
      if (wb.restoreDrawing(sheet, snapshot) === null) {
        throw new Error('structure: failed to restore a deleted drawing');
      }
    }
    for (const restore of deletedObjects) {
      if (!restore()) throw new Error('structure: failed to restore a deleted drawing');
    }
    if (deletedDrawings.length > 0) hydrateRestoredCharts(wb, store, sheet, deletedIds);
    hydrateRestoredImages(wb, store, sheet, deletedIds);
    if (deletedIds.size > 0) {
      mutators.setDrawingOrder(store, orderBefore);
      if (!writeDrawingOrder(wb, store.getState(), sheet)) {
        throw new Error('structure: failed to restore the drawing order');
      }
    }
    hydrateSlicersFromEngine(wb, store);
  };

  const rollbackRedo = (): void => {
    if (!runNativeOp(!positive)) throw new Error('structure: native redo rollback failed');
    if (canRestoreMerges) restoreMerges(beforeMerges);
    restoreDeletedBand();
    restoreCheckpoint(beforeAxes, beforeNative, beforeTracking, beforeRangeStyle);
    wb.recalcAuto();
  };

  const rollbackUndo = (): void => {
    if (!runNativeOp(positive)) throw new Error('structure: native undo rollback failed');
    if (canRestoreMerges) restoreMerges(afterMerges);
    restoreCheckpoint(afterAxes, afterNative, afterTracking, afterRangeStyle);
    wb.recalcAuto();
  };

  const apply = (): void => {
    let applied = false;
    try {
      if (!initialApply && replayCheckpoint) {
        replayCheckpoint.before =
          replayCheckpoint.capture?.() ?? captureStructureReplayCheckpoint(store, wb, sheet);
      }
      if (!runNativeOp(positive)) throw new Error('structure: native axis edit failed');
      applied = true;
      if (canRestoreMerges) {
        if (initialApply) afterMerges = wb.getMerges(sheet);
        else restoreMerges(afterMerges);
      }
      if (!initialApply) {
        restoreCheckpoint(afterAxes, afterNative, afterTracking, afterRangeStyle);
        wb.recalcAuto();
      }
      initialApply = false;
    } catch (error) {
      if (!applied) {
        const recoveryErrors: unknown[] = [];
        try {
          restoreAfterFailedReplay?.('redo');
        } catch (rollbackError) {
          recoveryErrors.push(rollbackError);
        }
        if (recoveryErrors.length > 0)
          throw new AggregateError([error, ...recoveryErrors], 'structure: redo recovery failed', {
            cause: error,
          });
        throw error;
      }
      const recoveryErrors: unknown[] = [];
      try {
        rollbackRedo();
      } catch (rollbackError) {
        recoveryErrors.push(rollbackError);
      }
      try {
        restoreAfterFailedReplay?.('redo');
      } catch (restoreError) {
        recoveryErrors.push(restoreError);
      }
      if (recoveryErrors.length > 0)
        throw new AggregateError([error, ...recoveryErrors], 'structure: redo rollback failed', {
          cause: error,
        });
      throw error;
    }
  };
  const invert = (): void => {
    let applied = false;
    try {
      if (replayCheckpoint?.after === undefined) {
        if (replayCheckpoint) {
          replayCheckpoint.after =
            replayCheckpoint.capture?.() ?? captureStructureReplayCheckpoint(store, wb, sheet);
        }
      }
      if (!runNativeOp(!positive)) throw new Error('structure: native axis undo failed');
      applied = true;
      if (canRestoreMerges) restoreMerges(beforeMerges);
      restoreDeletedBand();
      wb.recalcAuto();
      restoreCheckpoint(beforeAxes, beforeNative, beforeTracking, beforeRangeStyle);
      wb.recalcAuto();
    } catch (error) {
      if (!applied) {
        const recoveryErrors: unknown[] = [];
        try {
          restoreAfterFailedReplay?.('undo');
        } catch (rollbackError) {
          recoveryErrors.push(rollbackError);
        }
        if (recoveryErrors.length > 0)
          throw new AggregateError([error, ...recoveryErrors], 'structure: undo recovery failed', {
            cause: error,
          });
        throw error;
      }
      const recoveryErrors: unknown[] = [];
      try {
        rollbackUndo();
      } catch (rollbackError) {
        recoveryErrors.push(rollbackError);
      }
      try {
        restoreAfterFailedReplay?.('undo');
      } catch (restoreError) {
        recoveryErrors.push(restoreError);
      }
      if (recoveryErrors.length > 0)
        throw new AggregateError([error, ...recoveryErrors], 'structure: undo rollback failed', {
          cause: error,
        });
      throw error;
    }
  };

  let initialApplied = false;
  try {
    if (!runNativeOp(positive)) return false;
    initialApplied = true;
    if (canRestoreMerges) afterMerges = wb.getMerges(sheet);
    if (record) {
      afterAxes = !useRangeStyleCheckpoint ? axisSnapshotFor(wb, sheet) : null;
      if (useRangeStyleCheckpoint) {
        afterRangeStyle = wb.snapshotRangeStyle(styleRange);
        if (afterRangeStyle === null)
          throw new Error('structure: native range style after-snapshot failed');
      }
      if (bridge) {
        const keys = afterKeys();
        if (!useRangeStyleCheckpoint) afterNative = bridge.capture(keys);
        afterTracking = bridge.captureTracking(keys);
      }
    }
  } catch (error) {
    if (!initialApplied) throw error;
    const rollbackErrors: unknown[] = [];
    try {
      if (!runNativeOp(!positive)) throw new Error('structure: native axis rollback failed');
      if (canRestoreMerges) restoreMerges(beforeMerges);
      restoreDeletedBand();
      restoreCheckpoint(beforeAxes, beforeNative, beforeTracking, beforeRangeStyle);
      wb.recalcAuto();
    } catch (rollbackError) {
      rollbackErrors.push(rollbackError);
    }
    if (rollbackErrors.length > 0) {
      throw new AggregateError(
        [error, ...rollbackErrors],
        'structure: native axis edit rollback failed',
        {
          cause: error,
        },
      );
    }
    throw error;
  }
  initialApply = false;
  if (record) {
    history.push({ undo: invert, redo: apply });
  }
  return true;
}
