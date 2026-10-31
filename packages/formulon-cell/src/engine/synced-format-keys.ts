import { addrKey, parseAddrKey } from './address.js';
import { type AxisStructureShift, mapAxisAddress } from './axis-structure-remap.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * High-water mark of cells we have assigned a non-default XF to, per workbook
 * and sheet. Lets a later sync reset the XF of a cell whose format entry was
 * removed (Clear Formats) back to 0 — otherwise the engine keeps the stale XF
 * and the cleared format resurrects on the next save.
 */
const syncedFormatKeys = new WeakMap<WorkbookHandle, Map<number, Set<string>>>();

export function formattedKeySet(wb: WorkbookHandle, sheet: number): Set<string> {
  let perSheet = syncedFormatKeys.get(wb);
  if (!perSheet) {
    perSheet = new Map();
    syncedFormatKeys.set(wb, perSheet);
  }
  let set = perSheet.get(sheet);
  if (!set) {
    set = new Set();
    perSheet.set(sheet, set);
  }
  return set;
}

/** Record that `keys` currently carry a non-default XF on `sheet`, so a later
 *  sync knows to reset any that disappear. Used both after a writeback and
 *  after hydrating XFs from a loaded workbook. */
export function seedSyncedFormatKeys(
  wb: WorkbookHandle,
  sheet: number,
  keys: Iterable<string>,
): void {
  const set = formattedKeySet(wb, sheet);
  for (const key of keys) set.add(key);
}

/**
 * Capture only the membership bits for the supplied address keys. Format
 * history can temporarily flush a store snapshot while the native engine is
 * being used as the source of truth; on a failed flush it must restore the
 * high-water bookkeeping that existed before that attempt. Do not call
 * `formattedKeySet` here: a read-only checkpoint must not create a cache
 * bucket for an untouched workbook or sheet.
 */
export function captureSyncedFormatKeyMembership(
  wb: WorkbookHandle,
  keys: Iterable<string>,
): Map<string, boolean> {
  const snapshot = new Map<string, boolean>();
  const perSheet = syncedFormatKeys.get(wb);
  for (const key of keys) {
    const addr = parseAddrKey(key);
    if (!addr || addrKey(addr) !== key) continue;
    snapshot.set(key, perSheet?.get(addr.sheet)?.has(key) === true);
  }
  return snapshot;
}

/** Restore membership bits captured by `captureSyncedFormatKeyMembership`.
 * The supplied keys are the complete scope of the checkpoint: unrelated
 * sheets and cells retain their current tracking state. Empty buckets are
 * removed so this rollback does not leave inert WeakMap state behind. */
export function restoreSyncedFormatKeyMembership(
  wb: WorkbookHandle,
  snapshot: ReadonlyMap<string, boolean>,
): void {
  let perSheet = syncedFormatKeys.get(wb);
  for (const [key, tracked] of snapshot) {
    const addr = parseAddrKey(key);
    if (!addr || addrKey(addr) !== key) continue;

    if (tracked) {
      if (!perSheet) {
        perSheet = new Map();
        syncedFormatKeys.set(wb, perSheet);
      }
      let set = perSheet.get(addr.sheet);
      if (!set) {
        set = new Set();
        perSheet.set(addr.sheet, set);
      }
      set.add(key);
      continue;
    }

    const set = perSheet?.get(addr.sheet);
    if (!set) continue;
    set.delete(key);
    if (set.size === 0) perSheet?.delete(addr.sheet);
  }
  if (perSheet?.size === 0) syncedFormatKeys.delete(wb);
}

/** Keep the high-water reset cache aligned with a native sheet move/remove.
 * This must run before the corresponding store publication; otherwise a flush
 * of the newly mapped format entries can reset an unrelated sheet using the
 * old cache key. */
export function remapSyncedFormatKeys(
  wb: WorkbookHandle,
  mapIndex: (index: number) => number | null,
): void {
  const current = syncedFormatKeys.get(wb);
  if (!current) return;
  const next = new Map<number, Set<string>>();
  for (const [sheet, keys] of current) {
    const mappedSheet = mapIndex(sheet);
    if (mappedSheet === null) continue;
    let mappedKeys = next.get(mappedSheet);
    if (!mappedKeys) {
      mappedKeys = new Set<string>();
      next.set(mappedSheet, mappedKeys);
    }
    for (const key of keys) {
      const addr = parseAddrKey(key);
      // The outer sheet bucket and the address prefix must agree. A malformed
      // cache entry cannot safely be reset after a permutation, so discard it
      // rather than allowing a future flush to target an unrelated cell.
      if (!addr || addr.sheet !== sheet) continue;
      const mappedKey = mapIndex(addr.sheet);
      if (mappedKey !== null) mappedKeys.add(addrKey({ ...addr, sheet: mappedKey }));
    }
  }
  syncedFormatKeys.set(wb, next);
}

/** Keep the cell-XF high-water cache aligned with a row/column structure edit.
 * Only the selected sheet is remapped; all other sheet buckets retain their
 * current membership. Keys in a deleted band disappear from the cache. */
export function remapSyncedFormatKeysForStructure(
  wb: WorkbookHandle,
  sheet: number,
  shift: AxisStructureShift,
): void {
  const current = syncedFormatKeys.get(wb);
  const keys = current?.get(sheet);
  if (!current || !keys) return;
  const next = new Map(current);
  const mapped = new Set<string>();
  for (const key of keys) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const nextAddr = mapAxisAddress(addr, shift);
    if (nextAddr) mapped.add(addrKey(nextAddr));
  }
  if (mapped.size === 0) next.delete(sheet);
  else next.set(sheet, mapped);
  if (next.size === 0) syncedFormatKeys.delete(wb);
  else syncedFormatKeys.set(wb, next);
}

/** Capture the exact high-water membership for one sheet before a row/column
 * structure edit. `null` distinguishes an absent bucket from an empty one so
 * history replay can restore the cache without coordinate re-expansion. */
export function captureSyncedFormatKeysForStructure(
  wb: WorkbookHandle,
  sheet: number,
): Set<string> | null {
  const keys = syncedFormatKeys.get(wb)?.get(sheet);
  return keys ? new Set(keys) : null;
}

/** Restore a structure checkpoint's exact high-water membership. */
export function restoreSyncedFormatKeysForStructure(
  wb: WorkbookHandle,
  sheet: number,
  snapshot: ReadonlySet<string> | null,
): void {
  const current = syncedFormatKeys.get(wb);
  const next = current ? new Map(current) : new Map<number, Set<string>>();
  if (snapshot === null || snapshot.size === 0) next.delete(sheet);
  else next.set(sheet, new Set(snapshot));
  if (next.size === 0) syncedFormatKeys.delete(wb);
  else syncedFormatKeys.set(wb, next);
}
