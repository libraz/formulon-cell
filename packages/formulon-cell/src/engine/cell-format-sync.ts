import { sameStructuredValue } from '../commands/prepared-change.js';
import type { CellFormat, SheetAxisLocks, SpreadsheetStore } from '../store/store.js';
import { addrKey, parseAddrKey } from './address.js';
import { syncAxisFormatOverrides } from './axis-format-sync.js';
import { cellFormatFromXfWithStatus } from './cell-format-projection.js';
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
import { strictSyncError } from './strict-sync-error.js';
import { formattedKeySet, seedSyncedFormatKeys } from './synced-format-keys.js';
import type { Addr, CellRect, CellValue } from './types.js';
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

const strictError = (operation: string, address: string): Error =>
  strictSyncError('cell-format', operation, address);

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
    throw strictError('resolveStyleXf', `sheet:${sheet}`);
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
      if (strict) throw strictError('resolveXf', key);
      continue;
    }
    const setXfSucceeded = wb.setCellXfIndex(sheet, row, col, xfIndex);
    if (strict && !setXfSucceeded) {
      throw strictError('setCellXfIndex', key);
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
        throw strictError('setCellPhoneticRuns', key);
      }
    } else if (wb.capabilities.phonetic) {
      const setPhoneticSucceeded = wb.setCellPhonetic(
        sheet,
        row,
        col,
        flattenPhoneticRuns(fmt.phonetic),
      );
      if (strict && !setPhoneticSucceeded) {
        throw strictError('setCellPhonetic', key);
      }
    }
    if (!writePhoneticProperties(wb, { sheet, row, col }, fmt) && strict) {
      throw strictError('setCellPhoneticProperties', key);
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
      throw strictError('setCellShowPhonetic', key);
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
      throw strictError('resetCellXfIndex', key);
    }
  }
  strictStylePlan?.commit();
  previous.clear();
  for (const key of current) previous.add(key);
}

/** Engine surface that states each CellFormat field on hydrate; `null` marks a
 *  field held by the store or by another sync. Exhaustive over CellFormat, so a
 *  new field does not compile until it is classified here. */
type FormatFieldSource =
  | 'xf'
  | 'cellStyle'
  | 'phonetic'
  | 'phoneticProperties'
  | 'showPhonetic'
  | null;

const FORMAT_FIELD_SOURCES = {
  cellStyle: 'cellStyle',
  numFmt: 'xf',
  bold: 'xf',
  italic: 'xf',
  underline: 'xf',
  strike: 'xf',
  fontVertAlign: 'xf',
  fontOutline: 'xf',
  fontShadow: 'xf',
  fontCondense: 'xf',
  fontExtend: 'xf',
  phonetic: 'phonetic',
  phoneticProperties: 'phoneticProperties',
  showPhonetic: 'showPhonetic',
  align: 'xf',
  vAlign: 'xf',
  wrap: 'xf',
  justifyLastLine: 'xf',
  shrinkToFit: 'xf',
  indent: 'xf',
  rotation: 'xf',
  textDirection: 'xf',
  borders: 'xf',
  color: 'xf',
  fill: 'xf',
  fillPattern: 'xf',
  fillPatternColor: 'xf',
  fillGradient: 'xf',
  fontFamily: 'xf',
  fontScheme: 'xf',
  fontSize: 'xf',
  hyperlink: null,
  hyperlinkDisplay: null,
  hyperlinkTooltip: null,
  hyperlinkLocation: null,
  hyperlinkSpan: null,
  comment: null,
  commentAuthor: null,
  commentVisible: null,
  validation: null,
  locked: 'xf',
  formulaHidden: 'xf',
} as const satisfies Record<keyof CellFormat, FormatFieldSource>;

/** Fields this engine states on hydrate; a surface it cannot read is left to the store. */
function engineModelledFormatKeys(wb: WorkbookHandle): readonly (keyof CellFormat)[] {
  const caps = wb.capabilities;
  const phonetic = caps.phonetic === true || caps.phoneticRuns === true;
  const live: Record<NonNullable<FormatFieldSource>, boolean> = {
    xf: true,
    cellStyle: caps.cellStyles === true,
    phonetic,
    phoneticProperties: phonetic && caps.phoneticProperties === true,
    showPhonetic: caps.showPhonetic === true,
  };
  return (Object.keys(FORMAT_FIELD_SOURCES) as (keyof CellFormat)[]).filter((key) => {
    const source = FORMAT_FIELD_SOURCES[key];
    return source !== null && live[source];
  });
}

/**
 * Hydrate FormatSlice from engine XF entries on `sheet`, within `range` when
 * given. For each cell in scope the fields the engine models are replaced by
 * what it states now — set when stated, removed when not, including cells the
 * engine no longer lists — while fields it does not model (cell-level
 * `validation`, `comment`, `hyperlink` from the dedicated syncs) are kept. An
 * entry left with no fields is removed. A cell whose XF records cannot be read
 * in full is only merged into, never cleared.
 *
 * Skipped entirely when `capabilities.cellFormatting` is off. Workbook default
 * font/size are stripped so the store does not hold defaults the renderer
 * would already show.
 */
export function hydrateCellFormatsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  range?: CellRect,
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
  const inScope = (addr: Addr): boolean =>
    !range ||
    (addr.row >= range.firstRow &&
      addr.row <= range.lastRow &&
      addr.col >= range.firstCol &&
      addr.col <= range.lastCol);
  // What the engine states per listed cell; `mergeOnly` cells could not be read in full.
  const stated = new Map<string, Partial<CellFormat>>();
  const mergeOnly = new Set<string>();
  let axes: SheetAxisLocks | null | undefined;
  const axisLocks = (): SheetAxisLocks | null => {
    if (axes === undefined) axes = axisLocksFromEngine(wb, sheet);
    return axes;
  };
  // Styled blanks are absent from the physical cells; visit each address once.
  const xfByKey = new Map<string, { addr: Addr; value: CellValue; xfIndex: number }>();
  for (const { addr, xfIndex } of wb.cellXfIndices(sheet)) {
    if (inScope(addr)) xfByKey.set(addrKey(addr), { addr, value: { kind: 'blank' }, xfIndex });
  }
  for (const c of wb.physicalCells ? wb.physicalCells(sheet) : wb.cells(sheet)) {
    if (!inScope(c.addr)) continue;
    const key = addrKey(c.addr);
    xfByKey.set(key, { addr: c.addr, value: c.value, xfIndex: xfByKey.get(key)?.xfIndex ?? 0 });
  }
  // A blank whose only state is the Show Phonetic flag is in neither list above.
  for (const addr of wb.showPhoneticCells(sheet)) {
    const key = addrKey(addr);
    if (inScope(addr) && !xfByKey.has(key)) {
      xfByKey.set(key, { addr, value: { kind: 'blank' }, xfIndex: 0 });
    }
  }
  for (const c of xfByKey.values()) {
    const key = addrKey(c.addr);
    const phonetic = readPhoneticRuns(wb, c.addr, c.value);
    const xfIndex = c.xfIndex;
    if (xfIndex === null || xfIndex <= 0) {
      const patch = withShowPhonetic(wb, c.addr, phonetic ? { phonetic } : {});
      if (phonetic) withPhoneticProperties(wb, c.addr, patch);
      stated.set(key, patch);
      continue;
    }
    const xf = wb.getCellXf(xfIndex);
    if (!xf) {
      stated.set(key, {});
      mergeOnly.add(key);
      continue;
    }
    const projection = cellFormatFromXfWithStatus(wb, xf, workbookDefaultFont, styleKeys);
    const patch = projection.format;
    if (!projection.complete) mergeOnly.add(key);
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
    stated.set(key, patch);
  }
  // A report's look comes from its PivotTable style at paint time; only the
  // value fields' number formats land on the cells.
  if (wb.pivotCells) {
    for (const c of wb.pivotCells(sheet)) {
      if (!inScope(c.addr)) continue;
      const patch = pivotFormatPatch(c.numberFormat);
      if (Object.keys(patch).length === 0) continue;
      const key = addrKey(c.addr);
      stated.set(key, { ...stated.get(key), ...patch });
    }
  }
  const statedKeys: string[] = [];
  for (const [key, patch] of stated) {
    if (Object.keys(patch).length > 0) statedKeys.push(key);
  }
  // Seed the high-water mark so that clearing a format that was loaded from the
  // workbook (not authored in-session) still resets the engine XF.
  if (statedKeys.length > 0) seedSyncedFormatKeys(wb, sheet, statedKeys);

  const modelled = engineModelledFormatKeys(wb);
  const hasModelled = (format: CellFormat): boolean =>
    modelled.some((k) => Object.hasOwn(format, k));
  const current = store.getState().format.formats;
  const changes = new Map<string, CellFormat | null>();
  // One pass over the store: entries in scope the engine no longer lists lose its fields.
  const sheetPrefix = `${sheet}:`;
  for (const [key, prev] of current) {
    if (stated.has(key) || !key.startsWith(sheetPrefix)) continue;
    const addr = parseAddrKey(key);
    if (!addr || !inScope(addr) || !hasModelled(prev)) continue;
    changes.set(key, replaceModelledFields(prev, {}, modelled));
  }
  for (const [key, patch] of stated) {
    const prev = current.get(key);
    const hasPatch = Object.keys(patch).length > 0;
    if (mergeOnly.has(key)) {
      if (hasPatch) changes.set(key, { ...prev, ...patch });
    } else if (hasPatch || (prev && hasModelled(prev))) {
      changes.set(key, replaceModelledFields(prev ?? {}, patch, modelled));
    }
  }
  if (changes.size === 0) return;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const [key, next] of changes) {
      if (next) formats.set(key, next);
      else formats.delete(key);
    }
    return { ...s, format: { ...s.format, formats } };
  });
}

/** `prev` with every engine-modelled field replaced by `patch`; null when nothing is left. */
function replaceModelledFields(
  prev: CellFormat,
  patch: Partial<CellFormat>,
  modelled: readonly (keyof CellFormat)[],
): CellFormat | null {
  const next: CellFormat = { ...prev };
  for (const key of modelled) delete next[key];
  Object.assign(next, patch);
  return Object.keys(next).length > 0 ? next : null;
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
    throw strictError('sheetFontStyles', `sheet:${sheet}`);
  }
  const axesOk = syncAxisFormatOverrides(wb, store, [sheet]);
  if (options?.strict && !axesOk) {
    throw strictError('axisFormatStyles', `sheet:${sheet}`);
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
