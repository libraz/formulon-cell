import {
  type ActiveAxisFormatSnapshot,
  captureActiveAxisFormatOverrideForStructure,
  remapActiveAxisFormatOverridesForStructure,
  restoreActiveAxisFormatOverrideForStructure,
} from '../engine/axis-format-sync.js';
import type { AxisStructureShift } from '../engine/axis-structure-remap.js';
import {
  captureSyncedFormatKeysForStructure,
  remapSyncedFormatKeysForStructure,
  restoreSyncedFormatKeysForStructure,
} from '../engine/synced-format-keys.js';
import type { SheetAxisFormatOverride, SheetFontOverride } from '../engine/types.js';
import {
  type ActiveWholeSheetFontSnapshot,
  captureActiveWholeSheetFontOverrideForStructure,
  remapActiveWholeSheetFontOverridesForStructure,
  restoreActiveWholeSheetFontOverrideForStructure,
} from '../engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';

export const cloneAxisOverride = (
  override: SheetAxisFormatOverride | undefined,
): SheetAxisFormatOverride | undefined =>
  override
    ? {
        baseline: {
          columns: override.baseline.columns.map((entry) => ({ ...entry })),
          rows: override.baseline.rows.map((entry) => ({ ...entry })),
        },
        desired: {
          columns: override.desired.columns.map((entry) => ({ ...entry })),
          rows: override.desired.rows.map((entry) => ({ ...entry })),
        },
        ...(override.directXfs
          ? {
              directXfs: override.directXfs.map((entry) => ({
                key: entry.key,
                before: { ...entry.before },
                afterXfIndex: entry.afterXfIndex,
              })),
            }
          : {}),
        ...(override.fontPatch ? { fontPatch: { ...override.fontPatch } } : {}),
      }
    : undefined;

export const cloneSheetFontOverride = (
  override: SheetFontOverride | undefined,
): SheetFontOverride | undefined =>
  override
    ? {
        ...override,
        ...(override.patch ? { patch: { ...override.patch } } : {}),
        ...(override.baseline
          ? {
              baseline: {
                columns: override.baseline.columns.map((entry) => ({ ...entry })),
                rows: override.baseline.rows.map((entry) => ({ ...entry })),
              },
            }
          : {}),
      }
    : undefined;

export interface StructureCacheSnapshot {
  readonly activeAxis: ActiveAxisFormatSnapshot | null;
  readonly activeWholeSheetFont: ActiveWholeSheetFontSnapshot | null;
  readonly syncedFormatKeys: ReadonlySet<string> | null;
}

/** The native structure child is the first entry in a composite history
 * transaction. Later entries can already have restored their own before
 * snapshots when the native replay fails, so keep an isolated full store
 * checkpoint for putting those children back at the committed after-state. */
type StructureStoreSnapshot = State;

export interface StructureReplayCheckpoint {
  readonly store: StructureStoreSnapshot;
  readonly caches: StructureCacheSnapshot;
}

export interface StructureReplayCheckpointRef {
  readonly capture?: () => StructureReplayCheckpoint;
  before?: StructureReplayCheckpoint;
  after?: StructureReplayCheckpoint;
}

const captureStructureStoreSnapshot = (store: SpreadsheetStore): StructureStoreSnapshot =>
  structuredClone(store.getState()) as StructureStoreSnapshot;

export const restoreStructureStoreSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  snapshot: StructureStoreSnapshot,
  caches: StructureCacheSnapshot | undefined,
): void => {
  withEngineSyncMuted(wb, () => {
    store.setState(() => structuredClone(snapshot) as State);
    if (!caches) return;
    restoreActiveAxisFormatOverrideForStructure(wb, sheet, caches.activeAxis);
    restoreActiveWholeSheetFontOverrideForStructure(wb, sheet, caches.activeWholeSheetFont);
    restoreSyncedFormatKeysForStructure(wb, sheet, caches.syncedFormatKeys);
  });
};

export const captureStructureReplayCheckpoint = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
): StructureReplayCheckpoint => ({
  store: captureStructureStoreSnapshot(store),
  caches: {
    activeAxis: captureActiveAxisFormatOverrideForStructure(wb, sheet),
    activeWholeSheetFont: captureActiveWholeSheetFontOverrideForStructure(wb, sheet),
    syncedFormatKeys: captureSyncedFormatKeysForStructure(wb, sheet),
  },
});

/** Create a replay checkpoint handle for a composite command that invokes a
 * structural edit and appends more history children afterwards. The caller
 * owns the final marker; the structure child uses the same handle so native
 * failure recovery sees the committed outer transaction state. */
export const createStructureReplayCheckpointRef = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
): StructureReplayCheckpointRef => ({
  capture: () => captureStructureReplayCheckpoint(store, wb, sheet),
});

export const withEngineSyncMuted = <T>(wb: WorkbookHandle, fn: () => T): T => {
  const candidate = wb as WorkbookHandle & {
    withEngineSyncMuted?: <R>(callback: () => R) => R;
  };
  return typeof candidate.withEngineSyncMuted === 'function'
    ? candidate.withEngineSyncMuted(fn)
    : fn();
};

/** Publish a structural axis remap while the store mirror is muted. The
 * native operation has already shifted its records; this keeps all compact
 * JS/native caches in the same coordinate space before an ordinary store
 * subscriber can flush stale keys. */
export function applyStructureFormatRemap(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  shift: AxisStructureShift,
  axisOverride: SheetAxisFormatOverride | undefined,
  sheetFontOverride: SheetFontOverride | undefined,
  cacheSnapshot?: StructureCacheSnapshot,
): void {
  withEngineSyncMuted(wb, () => {
    store.setState((state) => {
      const currentAxes = state.format.sheetAxisOverrides;
      const nextAxes = currentAxes ? new Map(currentAxes) : new Map();
      const nextAxisOverride = cloneAxisOverride(axisOverride);
      if (nextAxisOverride) nextAxes.set(sheet, nextAxisOverride);
      else nextAxes.delete(sheet);
      const currentFonts = state.format.sheetFontOverrides;
      const nextFonts = currentFonts ? new Map(currentFonts) : new Map();
      const nextSheetFontOverride = cloneSheetFontOverride(sheetFontOverride);
      if (nextSheetFontOverride) nextFonts.set(sheet, nextSheetFontOverride);
      else nextFonts.delete(sheet);
      return {
        ...state,
        format: {
          ...state.format,
          sheetAxisOverrides: nextAxes,
          sheetFontOverrides: nextFonts,
        },
      };
    });
    if (cacheSnapshot) {
      restoreActiveAxisFormatOverrideForStructure(wb, sheet, cacheSnapshot.activeAxis);
      restoreActiveWholeSheetFontOverrideForStructure(
        wb,
        sheet,
        cacheSnapshot.activeWholeSheetFont,
      );
      restoreSyncedFormatKeysForStructure(wb, sheet, cacheSnapshot.syncedFormatKeys);
    } else {
      remapActiveAxisFormatOverridesForStructure(wb, sheet, shift);
      remapActiveWholeSheetFontOverridesForStructure(wb, sheet, shift);
      remapSyncedFormatKeysForStructure(wb, sheet, shift);
    }
  });
}
