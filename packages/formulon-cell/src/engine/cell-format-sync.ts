import { sameStructuredValue } from '../commands/prepared-change.js';
import type { CellFormat, SheetAxisLocks, SpreadsheetStore } from '../store/store.js';
import { addrKey, parseAddrKey } from './address.js';
import { syncAxisFormatOverrides } from './axis-format-sync.js';
import { cellFormatFromXf } from './cell-format-projection.js';
import {
  cellStyleKeysByXfId,
  importedCustomCellStyles,
  planNamedCellStylesToEngine,
  resolveNamedStyleXfs,
  syncNamedCellStylesToEngine,
} from './cell-style-writeback.js';
import { resolveStyleXf, resolveXfForFormat } from './cell-xf-writeback.js';
import { syncHyperlinksToEngine } from './format-sync.js';
import { formatCodeToNumFmt } from './format-writeback.js';
import {
  boundPhoneticRuns,
  flattenPhoneticRuns,
  readPhoneticRuns,
  withPhoneticProperties,
  withShowPhonetic,
  writePhoneticProperties,
  writeShowPhonetic,
} from './phonetic-sync.js';
import { axisLocksFromEngine, inheritsUnlocked } from './protection-sync.js';
import { formattedKeySet, seedSyncedFormatKeys } from './synced-format-keys.js';
import type { Addr, CellValue } from './types.js';
import { syncValidationsToEngine } from './validation-sync.js';
import { sheetFontPatchFor, syncWholeSheetFontOverrides } from './whole-sheet-font-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

const MAX_RESOLVED_XF_CACHE_ENTRIES = 256;

interface ResolvedXfCacheEntry {
  readonly key: string;
  readonly format: CellFormat;
  readonly xfIndex: number;
}

/** Exclude values whose formatter can observe state JSON and structuredClone
 *  cannot represent. Descriptor reads avoid invoking user getters. */
function isPlainDataTree(value: unknown, ancestors = new WeakSet<object>()): boolean {
  if (value === null || value === undefined) return true;
  switch (typeof value) {
    case 'string':
    case 'number':
    case 'boolean':
      return true;
    case 'bigint':
    case 'function':
    case 'symbol':
      return false;
    case 'object':
      break;
    default:
      return false;
  }

  const object = value as object;
  if (ancestors.has(object)) return false;
  ancestors.add(object);
  try {
    const prototype = Object.getPrototypeOf(object);
    const keys = Reflect.ownKeys(object);
    if (Array.isArray(object)) {
      if (prototype !== Array.prototype) return false;
      const length = Object.getOwnPropertyDescriptor(object, 'length');
      if (!length || !('value' in length) || !Number.isInteger(length.value)) return false;
      const arrayLength = length.value as number;
      if (arrayLength < 0 || keys.length !== arrayLength + 1) return false;
      const indexKeys = new Set<string>();
      for (const key of keys) {
        if (typeof key !== 'string') return false;
        if (key === 'length') continue;
        const index = Number(key);
        if (!Number.isInteger(index) || index < 0 || String(index) !== key) return false;
        indexKeys.add(key);
      }
      if (indexKeys.size !== arrayLength) return false;
      for (let index = 0; index < arrayLength; index += 1) {
        const descriptor = Object.getOwnPropertyDescriptor(object, String(index));
        if (!descriptor?.enumerable || !('value' in descriptor)) return false;
        if (!isPlainDataTree(descriptor.value, ancestors)) return false;
      }
      return true;
    }
    if (prototype !== Object.prototype && prototype !== null) return false;
    for (const key of keys) {
      if (typeof key !== 'string') return false;
      const descriptor = Object.getOwnPropertyDescriptor(object, key);
      if (!descriptor?.enumerable || !('value' in descriptor)) return false;
      if (!isPlainDataTree(descriptor.value, ancestors)) return false;
    }
    return true;
  } catch {
    return false;
  } finally {
    ancestors.delete(object);
  }
}

export interface EngineSyncOptions {
  readonly strict?: boolean;
}

const strictSyncError = (operation: string, address: string): Error =>
  new Error(`Strict engine sync failed: cell-format ${operation} at ${address}`);

/**
 * Push every format entry on `sheet` from FormatSlice into the engine's XF
 * table. For each cell the writeback ensures a font / fill / border / numFmt
 * record exists (the engine dedups against existing rows), assembles an XF,
 * and pins the cell's `xfIndex`. No-op when `capabilities.cellFormatting`
 * is off.
 *
 * Cells whose entry was removed from the store since the last sync (Clear
 * Formats) are reset to xfIndex 0 (the workbook default) via a high-water
 * mark, so a cleared format does not survive into a subsequent save.
 */
export function syncCellFormatsToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  options?: EngineSyncOptions,
): void {
  if (!wb.capabilities.cellFormatting) return;
  const strict = options?.strict === true;
  const state = store.getState();
  const formats = state.format.formats;
  let styleXfConstructionFailed = false;
  const resolveStyleXfForSync = (format: Partial<CellFormat>): number => {
    const xfIndex = resolveStyleXf(wb, format);
    if (strict && xfIndex < 0) styleXfConstructionFailed = true;
    return xfIndex;
  };
  const strictStylePlan = strict
    ? planNamedCellStylesToEngine(wb, state, resolveStyleXfForSync)
    : null;
  const styleXfIds = strictStylePlan
    ? strictStylePlan.styleXfIds
    : syncNamedCellStylesToEngine(wb, state, resolveStyleXfForSync);
  const namedStyleXfs = resolveNamedStyleXfs(wb, state, styleXfIds);
  if (strict && styleXfConstructionFailed) {
    throw strictSyncError('resolveStyleXf', `sheet:${sheet}`);
  }
  const previous = formattedKeySet(wb, sheet);
  const current = new Set<string>();
  const sheetFontPatch = sheetFontPatchFor(state, sheet);
  const resolvedXfs = new Map<string, Set<ResolvedXfCacheEntry>>();
  const resolvedXfLru = new Map<ResolvedXfCacheEntry, true>();
  const cacheKeyFor = (format: CellFormat, sourceXfIndex: number | null): string | null => {
    if (sourceXfIndex === null || !Number.isInteger(sourceXfIndex) || sourceXfIndex < 0) {
      return null;
    }
    if (!isPlainDataTree(format)) return null;
    // Patternless accounting codes come from workbook locale templates the key does not capture.
    if (format.numFmt?.kind === 'accounting' && format.numFmt.pattern == null) return null;
    try {
      const serialized = JSON.stringify(format);
      return serialized === undefined ? null : `${sourceXfIndex}\u0000${serialized}`;
    } catch {
      return null;
    }
  };
  const cachedXfFor = (key: string | null, format: CellFormat): number | null => {
    if (key === null) return null;
    const bucket = resolvedXfs.get(key);
    if (!bucket) return null;
    for (const entry of bucket) {
      try {
        if (!sameStructuredValue(entry.format, format)) continue;
      } catch {
        continue;
      }
      resolvedXfLru.delete(entry);
      resolvedXfLru.set(entry, true);
      return entry.xfIndex;
    }
    return null;
  };
  const rememberXf = (key: string | null, format: CellFormat, xfIndex: number): void => {
    if (key === null || xfIndex < 0) return;
    let snapshot: CellFormat;
    try {
      snapshot = structuredClone(format);
    } catch {
      return;
    }
    const entry: ResolvedXfCacheEntry = { key, format: snapshot, xfIndex };
    let bucket = resolvedXfs.get(key);
    if (!bucket) {
      bucket = new Set();
      resolvedXfs.set(key, bucket);
    }
    bucket.add(entry);
    resolvedXfLru.set(entry, true);
    while (resolvedXfLru.size > MAX_RESOLVED_XF_CACHE_ENTRIES) {
      const oldest = resolvedXfLru.keys().next().value;
      if (!oldest) break;
      resolvedXfLru.delete(oldest);
      const oldestBucket = resolvedXfs.get(oldest.key);
      oldestBucket?.delete(oldest);
      if (oldestBucket?.size === 0) resolvedXfs.delete(oldest.key);
    }
  };
  for (const [key, fmt] of formats) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const { row, col } = addr;
    const directXfIndex = wb.getCellXfIndex(sheet, row, col);
    let sourceXfIndex = directXfIndex;
    if (directXfIndex === 0 && wb.capabilities.effectiveStyle === true) {
      const effective = wb.getEffectiveStyle(sheet, row, col);
      if (effective?.source === 1 || effective?.source === 2) {
        sourceXfIndex = effective.xfIndex;
      }
    }
    const cacheKey = cacheKeyFor(fmt, sourceXfIndex);
    let xfIndex = cachedXfFor(cacheKey, fmt);
    if (xfIndex === null) {
      const resolution = resolveXfForFormat(
        wb,
        fmt,
        styleXfIds,
        namedStyleXfs,
        strict,
        sourceXfIndex,
        sheetFontPatch,
      );
      xfIndex = resolution.xfIndex;
      if (resolution.cacheable) rememberXf(cacheKey, fmt, xfIndex);
    }
    if (xfIndex < 0) {
      if (strict) throw strictSyncError('resolveXf', key);
      continue;
    }
    const setXfSucceeded = wb.setCellXfIndex(sheet, row, col, xfIndex);
    if (strict && !setXfSucceeded) {
      throw strictSyncError('setCellXfIndex', key);
    }
    if (strict) previous.add(key);
    // Per-run first: `setCellPhonetic` spans the whole cell, so on a partially
    // annotated cell it would flatten every span into one reading.
    if (wb.capabilities.phoneticRuns) {
      const setPhoneticSucceeded = wb.setCellPhoneticRuns(
        sheet,
        row,
        col,
        boundPhoneticRuns(wb, { sheet, row, col }, fmt.phonetic),
      );
      if (strict && !setPhoneticSucceeded) {
        throw strictSyncError('setCellPhoneticRuns', key);
      }
    } else if (wb.capabilities.phonetic) {
      const setPhoneticSucceeded = wb.setCellPhonetic(
        sheet,
        row,
        col,
        flattenPhoneticRuns(fmt.phonetic),
      );
      if (strict && !setPhoneticSucceeded) {
        throw strictSyncError('setCellPhonetic', key);
      }
    }
    if (!writePhoneticProperties(wb, { sheet, row, col }, fmt) && strict) {
      throw strictSyncError('setCellPhoneticProperties', key);
    }
    if (
      !writeShowPhonetic(
        wb,
        { sheet, row, col },
        fmt,
        state.data.cells.get(key)?.formula != null,
      ) &&
      strict
    ) {
      throw strictSyncError('setCellShowPhonetic', key);
    }
    current.add(key);
  }
  // Reset cells that were formatted before but no longer are — the cleared
  // format must not linger in the engine XF table.
  for (const key of previous) {
    if (current.has(key)) continue;
    const addr = parseAddrKey(key);
    if (!addr) continue;
    const resetSucceeded = wb.setCellXfIndex(sheet, addr.row, addr.col, 0);
    if (strict && !resetSucceeded) {
      throw strictSyncError('resetCellXfIndex', key);
    }
  }
  strictStylePlan?.commit();
  previous.clear();
  for (const key of current) previous.add(key);
}

/**
 * Hydrate FormatSlice from engine XF entries on `sheet`. For every populated
 * cell, read its xfIndex, resolve to the underlying records, translate back
 * into CellFormat, and merge into the existing FormatSlice entry (preserving
 * any field the engine doesn't model — e.g. cell-level `validation`,
 * `comment`, `hyperlink` that the dedicated syncs already wrote).
 *
 * Skipped entirely when `capabilities.cellFormatting` is off, or when an XF
 * resolves to the default record (xfIndex 0 with all defaults). Workbook
 * default font/size are stripped so we do not pollute the store with
 * defaults that the renderer would already show.
 */
export function hydrateCellFormatsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.cellFormatting) return;
  const workbookDefaultFont = wb.workbookDefaultFont;
  const styleKeys = cellStyleKeysByXfId(wb);
  const known = new Set((store.getState().format.customCellStyles ?? []).map((s) => s.id));
  const imported = importedCustomCellStyles(wb).filter((style) => !known.has(style.id));
  if (imported.length > 0) {
    store.setState((s) => ({
      ...s,
      format: {
        ...s.format,
        customCellStyles: [...(s.format.customCellStyles ?? []), ...imported],
      },
    }));
  }
  const updates: Array<{ key: string; patch: Partial<CellFormat> }> = [];
  let axes: SheetAxisLocks | null | undefined;
  const axisLocks = (): SheetAxisLocks | null => {
    if (axes === undefined) axes = axisLocksFromEngine(wb, sheet);
    return axes;
  };
  // Styled blanks are absent from the physical cells; visit each address once.
  const xfByKey = new Map<string, { addr: Addr; value: CellValue; xfIndex: number }>();
  for (const { addr, xfIndex } of wb.cellXfIndices(sheet)) {
    xfByKey.set(addrKey(addr), { addr, value: { kind: 'blank' }, xfIndex });
  }
  for (const c of wb.physicalCells ? wb.physicalCells(sheet) : wb.cells(sheet)) {
    const key = addrKey(c.addr);
    xfByKey.set(key, { addr: c.addr, value: c.value, xfIndex: xfByKey.get(key)?.xfIndex ?? 0 });
  }
  // A blank whose only state is the Show Phonetic flag is in neither list above.
  for (const addr of wb.showPhoneticCells(sheet)) {
    const key = addrKey(addr);
    if (!xfByKey.has(key)) xfByKey.set(key, { addr, value: { kind: 'blank' }, xfIndex: 0 });
  }
  for (const c of xfByKey.values()) {
    const phonetic = readPhoneticRuns(wb, c.addr, c.value);
    const xfIndex = c.xfIndex;
    if (xfIndex === null || xfIndex <= 0) {
      const patch = withShowPhonetic(wb, c.addr, phonetic ? { phonetic } : {});
      if (phonetic) withPhoneticProperties(wb, c.addr, patch);
      if (Object.keys(patch).length > 0) updates.push({ key: addrKey(c.addr), patch });
      continue;
    }
    const xf = wb.getCellXf(xfIndex);
    if (!xf) continue;
    const patch = cellFormatFromXf(wb, xf, workbookDefaultFont, styleKeys);
    // A cell's own XF with no protection element is locked whatever its row or column style says;
    // where that style is unlocked, the entry says so explicitly instead of inheriting it.
    if (
      xf.locked === null &&
      xf.hidden === null &&
      inheritsUnlocked(axisLocks(), c.addr.row, c.addr.col)
    )
      patch.locked = true;
    if (phonetic) {
      patch.phonetic = phonetic;
      withPhoneticProperties(wb, c.addr, patch);
    }
    withShowPhonetic(wb, c.addr, patch);
    if (Object.keys(patch).length > 0) {
      updates.push({ key: addrKey(c.addr), patch });
    }
  }
  // A report's look comes from its PivotTable style at paint time; only the
  // value fields' number formats land on the cells.
  if (wb.pivotCells) {
    for (const c of wb.pivotCells(sheet)) {
      const patch = pivotFormatPatch(c.numberFormat);
      if (Object.keys(patch).length > 0) updates.push({ key: addrKey(c.addr), patch });
    }
  }
  if (updates.length === 0) return;
  // Seed the high-water mark so that clearing a format that was loaded from the
  // workbook (not authored in-session) still resets the engine XF.
  seedSyncedFormatKeys(
    wb,
    sheet,
    updates.map((u) => u.key),
  );
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const u of updates) {
      const prev = formats.get(u.key) ?? {};
      formats.set(u.key, { ...prev, ...u.patch });
    }
    return { ...s, format: { ...s.format, formats } };
  });
}

function pivotFormatPatch(numberFormat: string): Partial<CellFormat> {
  const patch: Partial<CellFormat> = {};
  if (numberFormat) {
    const numFmt = formatCodeToNumFmt(numberFormat);
    if (numFmt) patch.numFmt = numFmt;
  }
  return patch;
}

/** Publish the named-style catalog without touching any cell, for a style-list
 *  change on a store whose cells need no rewrite. */
export function syncNamedStyleCatalogToEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.cellFormatting) return;
  syncNamedCellStylesToEngine(wb, store.getState(), (format) => resolveStyleXf(wb, format));
}

/** One-shot flush of every store-side format dimension that has an engine
 *  surface: cell XF assignments, list-validation rules, and hyperlinks. Call
 *  after any format mutation that should round-trip through xlsx. Each
 *  per-dimension sync short-circuits on its own capability flag, so engines
 *  that only support a subset still work. */
export function flushFormatToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  options?: EngineSyncOptions,
): boolean {
  const sheetFontsOk = syncWholeSheetFontOverrides(wb, store, [sheet]);
  if (options?.strict && !sheetFontsOk) {
    throw strictSyncError('sheetFontStyles', `sheet:${sheet}`);
  }
  const axesOk = syncAxisFormatOverrides(wb, store, [sheet]);
  if (options?.strict && !axesOk) {
    throw strictSyncError('axisFormatStyles', `sheet:${sheet}`);
  }
  // Publish compact native row/column styles before sparse cell XFs. Axis
  // application can clear dimension records and must therefore be complete
  // before direct cells are written; strict callers also avoid a partial cell
  // publication when an axis setter rejects the transaction.
  syncCellFormatsToEngine(wb, store, sheet, options);
  syncValidationsToEngine(wb, store, sheet, options);
  syncHyperlinksToEngine(wb, store, sheet, options);
  return sheetFontsOk;
}
