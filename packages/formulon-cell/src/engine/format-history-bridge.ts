import type { SpreadsheetStore } from '../store/store.js';
import { parseAddrKey } from './address.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * The native state that a format history entry has to restore after the
 * store-side format has been flushed.  `directXfIndex` is the cell's own XF;
 * `effectiveXfIndex` is kept for cells whose formatting is inherited from a
 * row, column, or the workbook default.
 */
export interface NativeFormatSnapshotEntry {
  readonly directXfIndex: number | null;
  readonly effectiveXfIndex: number | null;
  /** Formulon's EffectiveStyleSource ordinal (Cell=0, Row=1, Column=2,
   * Default=3). Kept numeric here so this bridge stays compatible with both
   * the WASM and native npm type surfaces. */
  readonly source: number | null;
}

export type NativeFormatSnapshot = Map<string, NativeFormatSnapshotEntry>;

/** Membership of the format-sync high-water cache for the native keys in a
 * format history entry. The cache is deliberately opaque to this bridge: the
 * cell-format sync module owns its representation and exposes only this
 * transaction-safe membership view. */
export type FormatHistoryTrackingSnapshot = Map<string, boolean>;

/** Engine operations needed by the generic format-history recorder. */
export interface FormatHistoryBridge {
  readonly workbook: WorkbookHandle;
  withSyncMuted<T>(fn: () => T): T;
  capture(keys: Iterable<string>): NativeFormatSnapshot;
  captureTracking(keys: Iterable<string>): FormatHistoryTrackingSnapshot;
  flush(sheets: Iterable<number>): void;
  restore(snapshot: NativeFormatSnapshot): void;
  restoreTracking(snapshot: FormatHistoryTrackingSnapshot): void;
}

const bridges = new WeakMap<SpreadsheetStore, FormatHistoryBridge>();

/** Resolve the bridge for a mounted store, if that store currently owns one. */
export function formatHistoryBridgeFor(store: SpreadsheetStore): FormatHistoryBridge | undefined {
  return bridges.get(store);
}

/** Register the latest StoreMirror for a store. */
export function registerFormatHistoryBridge(
  store: SpreadsheetStore,
  bridge: FormatHistoryBridge,
): void {
  bridges.set(store, bridge);
}

/** Remove only the bridge that the caller installed. A stale mirror must not
 * remove a newer mirror's bridge when both temporarily share a store. */
export function unregisterFormatHistoryBridge(
  store: SpreadsheetStore,
  bridge: FormatHistoryBridge,
): void {
  if (bridges.get(store) === bridge) bridges.delete(store);
}

/** Build the store-bound bridge used by generic format history. */
export function createFormatHistoryBridge(
  handle: WorkbookHandle,
  flushSheet: (sheet: number) => void,
  captureTracking: (keys: Iterable<string>) => FormatHistoryTrackingSnapshot,
  restoreTracking: (snapshot: FormatHistoryTrackingSnapshot) => void,
): FormatHistoryBridge {
  return {
    workbook: handle,
    withSyncMuted: <T>(fn: () => T): T => handle.withEngineSyncMuted(fn),

    capture: (keys: Iterable<string>): NativeFormatSnapshot => {
      const snapshot: NativeFormatSnapshot = new Map();
      if (!handle.capabilities.cellFormatting) return snapshot;
      for (const key of keys) {
        const addr = parseAddrKey(key);
        if (!addr) continue;
        const directXfIndex = handle.getCellXfIndex(addr.sheet, addr.row, addr.col);
        if (directXfIndex === null) {
          throw new Error(`format history native XF snapshot unavailable at ${key}`);
        }
        const effective = handle.getEffectiveStyle(addr.sheet, addr.row, addr.col);
        if (handle.capabilities.effectiveStyle && effective === null) {
          throw new Error(`format history effective style snapshot unavailable at ${key}`);
        }
        snapshot.set(key, {
          directXfIndex,
          effectiveXfIndex: effective?.xfIndex ?? null,
          source: effective?.source ?? null,
        });
      }
      return snapshot;
    },

    captureTracking: (keys: Iterable<string>): FormatHistoryTrackingSnapshot =>
      captureTracking(keys),

    flush: (sheets: Iterable<number>): void => {
      for (const sheet of new Set(sheets)) {
        if (sheet < 0 || sheet >= handle.sheetCount) continue;
        flushSheet(sheet);
      }
    },

    restore: (snapshot: NativeFormatSnapshot): void => {
      const clearCellXfIndex = (
        handle as WorkbookHandle & {
          clearCellXfIndex?: (sheet: number, row: number, col: number) => boolean;
        }
      ).clearCellXfIndex;
      for (const [key, entry] of snapshot) {
        const addr = parseAddrKey(key);
        if (!addr) continue;
        const inherited = entry.source !== null && entry.source !== 0;
        const xfIndex = inherited ? entry.effectiveXfIndex : entry.directXfIndex;
        if (inherited && handle.capabilities.cellFormattingClear === true) {
          if (
            typeof clearCellXfIndex !== 'function' ||
            !clearCellXfIndex.call(handle, addr.sheet, addr.row, addr.col)
          ) {
            throw new Error(`format history native XF clear failed at ${key}`);
          }
          // The core can report a successful clear while retaining a
          // cell-owned XF for non-empty cells. Verify the inherited source;
          // only that semantic mismatch uses the effective-XF fallback. A
          // false clear result above remains a hard rollback error.
          const effective = handle.getEffectiveStyle(addr.sheet, addr.row, addr.col);
          if (
            effective?.source === entry.source &&
            (entry.effectiveXfIndex === null || effective.xfIndex === entry.effectiveXfIndex)
          ) {
            continue;
          }
          if (xfIndex === null || !handle.setCellXfIndex(addr.sheet, addr.row, addr.col, xfIndex)) {
            throw new Error(`format history native XF fallback restore failed at ${key}`);
          }
          continue;
        }
        if (xfIndex === null) continue;
        if (!handle.setCellXfIndex(addr.sheet, addr.row, addr.col, xfIndex)) {
          throw new Error(`format history native XF restore failed at ${key}`);
        }
      }
    },

    restoreTracking: (snapshot: FormatHistoryTrackingSnapshot): void => {
      restoreTracking(snapshot);
    },
  };
}
