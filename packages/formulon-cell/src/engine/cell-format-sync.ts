import type { CellStyleFormatGroup } from '../commands/cell-style-catalog.js';
import { CELL_STYLE_GROUP_FIELDS, cellStyleGroupMatches } from '../commands/cell-style-groups.js';
import {
  customPivotTableStyleById,
  pivotTableStyleAssignment,
  tableStyleSwatch,
} from '../commands/format-as-table.js';
import { withFillReplacement } from '../store/format-merge.js';
import { formatWithPending } from '../store/pending-format.js';
import type { CellFormat, SpreadsheetStore } from '../store/store.js';
import { addrKey, parseAddrKey } from './address.js';
import { syncAxisFormatOverrides } from './axis-format-sync.js';
import { type AxisStructureShift, mapAxisAddress } from './axis-structure-remap.js';
import {
  cellStyleKeysByXfId,
  importedCustomCellStyles,
  type NamedStyleXfResolution,
  planNamedCellStylesToEngine,
  resolveNamedStyleXfs,
  syncNamedCellStylesToEngine,
} from './cell-style-writeback.js';
import { fontCssPxToPoints, fontPointsToCssPx } from './font-units.js';
import { syncHyperlinksToEngine } from './format-sync.js';
import {
  BUILTIN_NUM_FMT_GENERAL,
  borderRecordFromFormat,
  borderRecordToFormat,
  buildXfRecord,
  cssColorToArgb,
  fillRecordFromFormat,
  fillRecordToFormat,
  fontRecordFromFormat,
  fontRecordToFormat,
  formatCodeToNumFmt,
  numFmtToFormatCode,
  readingOrderToTextDirection,
  textRotationToCellRotation,
} from './format-writeback.js';
import {
  withPhoneticProperties,
  withShowPhonetic,
  writePhoneticProperties,
  writeShowPhonetic,
} from './phonetic-sync.js';
import type {
  Addr,
  CellValue,
  CellXf,
  ColorSpec,
  FontRecord,
  PhoneticRun,
  SheetFontPatch,
} from './types.js';
import { syncValidationsToEngine } from './validation-sync.js';
import {
  formatWithSheetFontPatch,
  sheetFontPatchFor,
  syncWholeSheetFontOverrides,
} from './whole-sheet-font-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

const HALIGN_GENERAL = 0;
const VALIGN_BOTTOM = 2;

const PIVOT_KIND = {
  Header: 0,
  RowLabel: 1,
  ColLabel: 2,
  Data: 3,
  RowSubtotal: 4,
  ColSubtotal: 5,
  GrandTotal: 6,
} as const;

/**
 * High-water mark of cells we have assigned a non-default XF to, per workbook
 * and sheet. Lets a later sync reset the XF of a cell whose format entry was
 * removed (Clear Formats) back to 0 — otherwise the engine keeps the stale XF
 * and the cleared format resurrects on the next save.
 */
const syncedFormatKeys = new WeakMap<WorkbookHandle, Map<number, Set<string>>>();

export interface EngineSyncOptions {
  readonly strict?: boolean;
}

const strictSyncError = (operation: string, address: string): Error =>
  new Error(`Strict engine sync failed: cell-format ${operation} at ${address}`);

function formattedKeySet(wb: WorkbookHandle, sheet: number): Set<string> {
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

/** Concatenate a guide's readings into the single string the whole-cell
 *  phonetic entry points carry. */
export const flattenPhoneticRuns = (runs: readonly PhoneticRun[] | undefined): string =>
  runs === undefined ? '' : runs.map((run) => run.text).join('');

/**
 * Trim a guide to the spans that actually cover the cell's own text. A guide
 * travels with the format it belongs to, so pasting formats onto a shorter
 * cell would otherwise state runs reaching past the end of the string — the
 * engine takes those verbatim and writes them into the file, where no reader
 * can resolve them. A cell holding anything but text carries no guide at all.
 */
function boundPhoneticRuns(
  wb: WorkbookHandle,
  addr: Addr,
  runs: readonly PhoneticRun[] | undefined,
): readonly PhoneticRun[] {
  if (runs === undefined || runs.length === 0) return [];
  const value = wb.getValue(addr);
  const length = value.kind === 'text' ? value.value.length : 0;
  if (length === 0) return [];
  const bounded: PhoneticRun[] = [];
  for (const run of runs) {
    const start = Math.max(0, Math.min(run.start, length));
    const end = Math.max(start, Math.min(run.end, length));
    if (end === start) continue;
    bounded.push(start === run.start && end === run.end ? run : { start, end, text: run.text });
  }
  return bounded;
}

/** Read a cell's phonetic guide as runs, preferring the per-run surface. An
 *  engine that carries only the whole-cell reading reports one run spanning
 *  the cell text, which is how OOXML spells the same thing. */
function readPhoneticRuns(
  wb: WorkbookHandle,
  addr: Addr,
  value: CellValue,
): readonly PhoneticRun[] | null {
  if (wb.capabilities.phoneticRuns && typeof wb.getCellPhoneticRuns === 'function') {
    const runs = wb.getCellPhoneticRuns(addr.sheet, addr.row, addr.col);
    return runs !== null && runs.length > 0 ? runs : null;
  }
  if (wb.capabilities.phonetic && typeof wb.getCellPhonetic === 'function') {
    const text = wb.getCellPhonetic(addr.sheet, addr.row, addr.col);
    if (!text) return null;
    return [{ start: 0, end: value.kind === 'text' ? value.value.length : 0, text }];
  }
  return null;
}

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
  for (const [key, fmt] of formats) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const { row, col } = addr;
    const xfIndex = resolveXfForFormat(
      wb,
      fmt,
      styleXfIds,
      { sheet, row, col },
      namedStyleXfs,
      strict,
      sheetFontPatch,
    );
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
    if (wb.capabilities.phoneticRuns && typeof wb.setCellPhoneticRuns === 'function') {
      const setPhoneticSucceeded = wb.setCellPhoneticRuns(
        sheet,
        row,
        col,
        boundPhoneticRuns(wb, { sheet, row, col }, fmt.phonetic),
      );
      if (strict && !setPhoneticSucceeded) {
        throw strictSyncError('setCellPhoneticRuns', key);
      }
    } else if (wb.capabilities.phonetic && typeof wb.setCellPhonetic === 'function') {
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
  const physicalCells = wb.physicalCells ? wb.physicalCells(sheet) : wb.cells(sheet);
  for (const c of physicalCells) {
    const phonetic = readPhoneticRuns(wb, c.addr, c.value);
    const xfIndex = wb.getCellXfIndex(sheet, c.addr.row, c.addr.col);
    if (xfIndex === null || xfIndex <= 0) {
      const patch = withShowPhonetic(wb, c.addr, phonetic ? { phonetic } : {});
      if (phonetic) withPhoneticProperties(wb, c.addr, patch);
      if (Object.keys(patch).length > 0) updates.push({ key: addrKey(c.addr), patch });
      continue;
    }
    const xf = wb.getCellXf(xfIndex);
    if (!xf) continue;
    const patch = cellFormatFromXf(wb, xf, workbookDefaultFont, styleKeys);
    if (phonetic) {
      patch.phonetic = phonetic;
      withPhoneticProperties(wb, c.addr, patch);
    }
    withShowPhonetic(wb, c.addr, patch);
    if (Object.keys(patch).length > 0) {
      updates.push({ key: addrKey(c.addr), patch });
    }
  }
  if (wb.pivotCells) {
    for (const c of wb.pivotCells(sheet)) {
      const assignment =
        typeof c.pivotIndex === 'number'
          ? pivotTableStyleAssignment(store.getState(), sheet, c.pivotIndex)
          : null;
      const style = assignment
        ? customPivotTableStyleById(store.getState(), assignment.styleId)
        : null;
      const patch = pivotFormatPatch(c.kind, c.numberFormat, style);
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

function pivotFormatPatch(
  kind: number,
  numberFormat: string,
  style?: ReturnType<typeof customPivotTableStyleById>,
): Partial<CellFormat> {
  const patch: Partial<CellFormat> = {};
  const swatch = style ? tableStyleSwatch(style.style, style.color) : null;
  const headerFill = swatch?.header ?? '#d9eaf7';
  const bandFill = swatch?.band ?? '#eaf3f8';
  const totalFill = swatch ? swatch.header : '#bdd7ee';
  const headerText = swatch?.headerText ?? '#1f4e79';
  const blueRule = { style: 'thin' as const, color: swatch?.base ?? '#9dc3e6' };
  const lightRule = { style: 'thin' as const, color: swatch?.band ?? '#d9eaf7' };
  const totalRule = { style: 'medium' as const, color: swatch?.base ?? '#5b9bd5' };
  if (numberFormat) {
    const numFmt = formatCodeToNumFmt(numberFormat);
    if (numFmt) patch.numFmt = numFmt;
  }
  if (kind === PIVOT_KIND.Header || kind === PIVOT_KIND.RowLabel || kind === PIVOT_KIND.ColLabel) {
    patch.bold = true;
    patch.fill = headerFill;
    patch.color = headerText;
    patch.borders = { top: blueRule, bottom: blueRule };
  } else if (kind === PIVOT_KIND.RowSubtotal || kind === PIVOT_KIND.ColSubtotal) {
    patch.bold = true;
    patch.fill = bandFill;
    patch.borders = { top: lightRule, bottom: blueRule };
  } else if (kind === PIVOT_KIND.GrandTotal) {
    patch.bold = true;
    patch.fill = totalFill;
    patch.color = headerText;
    patch.borders = { top: totalRule, bottom: totalRule };
  }
  return patch;
}

export function cellFormatFromXf(
  wb: WorkbookHandle,
  xf: CellXf,
  workbookDefaultFont: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null = null,
  styleKeysByXfId: ReadonlyMap<number, string> | null = null,
): Partial<CellFormat> {
  const patch: Partial<CellFormat> = {};
  const styleKey = xf.xfId === undefined ? undefined : styleKeysByXfId?.get(xf.xfId);
  if (styleKey !== undefined) patch.cellStyle = styleKey;
  const resolved = wb as WorkbookHandle & {
    getResolvedFontRecord?: (fontIndex: number) => FontRecord | null;
    getResolvedFillRecord?: (fillIndex: number) => ReturnType<WorkbookHandle['getFillRecord']>;
    getResolvedBorderRecord?: (
      borderIndex: number,
    ) => ReturnType<WorkbookHandle['getBorderRecord']>;
  };
  const font =
    typeof resolved.getResolvedFontRecord === 'function'
      ? resolved.getResolvedFontRecord(xf.fontIndex)
      : wb.getFontRecord(xf.fontIndex);
  if (font) Object.assign(patch, fontRecordToFormat(font, workbookDefaultFont));
  const fill =
    typeof resolved.getResolvedFillRecord === 'function'
      ? resolved.getResolvedFillRecord(xf.fillIndex)
      : wb.getFillRecord(xf.fillIndex);
  if (fill) Object.assign(patch, fillRecordToFormat(fill));
  const border =
    typeof resolved.getResolvedBorderRecord === 'function'
      ? resolved.getResolvedBorderRecord(xf.borderIndex)
      : wb.getBorderRecord(xf.borderIndex);
  if (border) Object.assign(patch, borderRecordToFormat(border));
  if (xf.numFmtId !== BUILTIN_NUM_FMT_GENERAL) {
    const code = wb.getNumFmtCode(xf.numFmtId);
    if (code !== null) {
      const numFmt = formatCodeToNumFmt(code);
      if (numFmt) patch.numFmt = numFmt;
    }
  }
  const hasAlignmentChild = xf.hasAlignment !== false;
  const hasHorizontalAlign =
    hasAlignmentChild &&
    (xf.hasHorizontalAlign === true ||
      (xf.hasHorizontalAlign === undefined && xf.horizontalAlign !== HALIGN_GENERAL));
  const hasVerticalAlign =
    hasAlignmentChild &&
    (xf.hasVerticalAlign === true ||
      (xf.hasVerticalAlign === undefined && xf.verticalAlign !== VALIGN_BOTTOM));
  const hasWrapText =
    hasAlignmentChild && (xf.hasWrapText === true || (xf.hasWrapText === undefined && xf.wrapText));
  const hasJustifyLastLine =
    hasAlignmentChild &&
    (xf.hasJustifyLastLine === true ||
      (xf.hasJustifyLastLine === undefined && xf.justifyLastLine === true));
  if (hasHorizontalAlign && xf.horizontalAlign === 1) patch.align = 'left';
  else if (hasHorizontalAlign && xf.horizontalAlign === 2) patch.align = 'center';
  else if (hasHorizontalAlign && xf.horizontalAlign === 3) patch.align = 'right';
  else if (hasHorizontalAlign && xf.horizontalAlign === 4) patch.align = 'fill';
  else if (hasHorizontalAlign && xf.horizontalAlign === 5) patch.align = 'justify';
  else if (hasHorizontalAlign && xf.horizontalAlign === 6) patch.align = 'centerContinuous';
  else if (hasHorizontalAlign && xf.horizontalAlign === 7) patch.align = 'distributed';
  if (hasVerticalAlign && xf.verticalAlign === 0) patch.vAlign = 'top';
  else if (hasVerticalAlign && xf.verticalAlign === 1) patch.vAlign = 'middle';
  else if (hasVerticalAlign && xf.verticalAlign === 2) patch.vAlign = 'bottom';
  else if (hasVerticalAlign && xf.verticalAlign === 3) patch.vAlign = 'justify';
  else if (hasVerticalAlign && xf.verticalAlign === 4) patch.vAlign = 'distributed';
  if (hasWrapText) patch.wrap = xf.wrapText;
  if (hasJustifyLastLine) patch.justifyLastLine = xf.justifyLastLine === true;
  if (hasAlignmentChild) {
    const rotation = textRotationToCellRotation(xf.textRotation);
    if (rotation !== undefined) patch.rotation = rotation;
    if (xf.indent !== undefined) patch.indent = xf.indent;
    if (xf.shrinkToFit !== undefined) patch.shrinkToFit = xf.shrinkToFit;
    const direction = readingOrderToTextDirection(xf.readingOrder);
    if (direction !== undefined) patch.textDirection = direction;
  }
  if (xf.hasProtection === true) {
    patch.locked = xf.locked !== false;
    patch.formulaHidden = xf.hidden === true;
  }
  return patch;
}

/**
 * Project the engine's effective style for rendering a cell. The result is a
 * display projection only: when the engine selected a row or column XF, the
 * caller must keep the store entry absent so a later format flush does not
 * flatten inherited formatting into a direct cell XF.
 */
export function effectiveCellFormatFromEngine(
  wb: WorkbookHandle,
  sheet: number,
  row: number,
  col: number,
  cache?: Map<string, Partial<CellFormat> | null>,
): Partial<CellFormat> | null {
  if (!wb.capabilities?.effectiveStyle) {
    const displayFont = wb.workbookDisplayFont;
    return displayFont
      ? {
          fontFamily: displayFont.name,
          fontSize: fontPointsToCssPx(displayFont.size),
        }
      : null;
  }
  const effective = wb.getEffectiveStyle(sheet, row, col);
  if (!effective) return null;
  // The projection depends only on these fields, so a caller painting many
  // cells that share an XF can reuse one result; it must not mutate it.
  const cacheKey = `${effective.xfIndex}|${effective.locked}|${effective.hidden}|${effective.numFmtCode}`;
  const cached = cache?.get(cacheKey);
  if (cached !== undefined) return cached;
  const format = projectEffectiveStyle(wb, effective);
  cache?.set(cacheKey, format);
  return format;
}

function projectEffectiveStyle(
  wb: WorkbookHandle,
  effective: NonNullable<ReturnType<WorkbookHandle['getEffectiveStyle']>>,
): Partial<CellFormat> | null {
  const xf = wb.getCellXf(effective.xfIndex);
  if (!xf) return null;
  const format = cellFormatFromXf(wb, xf, wb.workbookDefaultFont);
  // `cellFormatFromXf` intentionally strips the workbook default from the
  // sparse store representation. Rendering still needs the actual workbook
  // font, including a theme-resolved family and its native point size.
  const displayFont = wb.workbookDisplayFont;
  if (displayFont) {
    if (!Object.hasOwn(format, 'fontFamily')) format.fontFamily = displayFont.name;
    if (!Object.hasOwn(format, 'fontSize')) {
      format.fontSize = fontPointsToCssPx(displayFont.size);
    }
  }
  // The effective result carries protection and the resolved number-format
  // code even when the selected XF omits an explicit protection/numFmt
  // element. Keep those fields in this paint-only projection; this helper is
  // never fed back into the direct XF synchronizer.
  if (!Object.hasOwn(format, 'locked')) format.locked = effective.locked;
  if (!Object.hasOwn(format, 'formulaHidden')) format.formulaHidden = effective.hidden;
  if (!Object.hasOwn(format, 'numFmt') && effective.numFmtCode) {
    const numFmt = formatCodeToNumFmt(effective.numFmtCode);
    if (numFmt) format.numFmt = numFmt;
  }
  return format;
}

/**
 * Display-only effective format for one address. Engine style (including row
 * and column inheritance) is the base; direct and pending store fields then
 * take precedence. The result must never be written back to the format slice.
 */
export function effectiveCellFormatWithPending(
  wb: WorkbookHandle,
  state: import('../store/store.js').State,
  addr: Addr,
): Partial<CellFormat> | null {
  const effective = effectiveCellFormatFromEngine(wb, addr.sheet, addr.row, addr.col);
  const direct = formatWithPending(state, addr);
  if (!effective && !direct) return null;
  return { ...(effective ?? {}), ...(direct ?? {}) };
}

/** A raw alignment group can only be reused when every authored extended
 *  field has a representable source value. Without this guard the normalised
 *  style matcher would treat absent/255/explicit-default values as equal and
 *  copy stale raw metadata over a deliberate reset. */
function rawAlignmentHasAuthoredDifference(fmt: CellFormat, raw: CellXf): boolean {
  if (raw.hasAlignment === false) {
    return (
      fmt.align !== undefined ||
      fmt.vAlign !== undefined ||
      fmt.wrap !== undefined ||
      fmt.justifyLastLine !== undefined ||
      fmt.rotation !== undefined ||
      fmt.indent !== undefined ||
      fmt.shrinkToFit !== undefined ||
      fmt.textDirection !== undefined
    );
  }
  if (fmt.align !== undefined && raw.hasHorizontalAlign !== true) return true;
  if (fmt.vAlign !== undefined && raw.hasVerticalAlign !== true) return true;
  if (fmt.wrap !== undefined && raw.hasWrapText !== true) return true;
  if (fmt.justifyLastLine !== undefined && raw.hasJustifyLastLine !== true) return true;
  if (fmt.rotation !== undefined) {
    if (raw.textRotation === undefined) return true;
    if (raw.textRotation === 255) return true;
    if (textRotationToCellRotation(raw.textRotation) !== fmt.rotation) return true;
  }
  if (fmt.indent !== undefined) {
    if (raw.indent === undefined || Math.round(fmt.indent) !== raw.indent) return true;
  }
  if (fmt.shrinkToFit !== undefined && raw.shrinkToFit !== fmt.shrinkToFit) return true;
  if (fmt.textDirection !== undefined) {
    if (raw.readingOrder === undefined) return true;
    if (readingOrderToTextDirection(raw.readingOrder) !== fmt.textDirection) return true;
  }
  return false;
}

/** Assemble the XF record for a CellFormat, ensuring every component record
 *  exists first (the engine dedups on add). Returns null on engine failure. */
function buildXfForFormat(
  wb: WorkbookHandle,
  fmt: CellFormat,
  sheetFontPatch?: SheetFontPatch,
): CellXf | null {
  const effectiveFormat = formatWithSheetFontPatch(fmt, sheetFontPatch);
  const fontIndex = wb.addFontRecord(fontRecordFromFormat(effectiveFormat, wb.workbookDefaultFont));
  if (fontIndex < 0) return null;
  const fillIndex = wb.addFillRecord(fillRecordFromFormat(fmt));
  if (fillIndex < 0) return null;
  const borderIndex = wb.addBorderRecord(borderRecordFromFormat(fmt));
  if (borderIndex < 0) return null;
  const code = numFmtToFormatCode(fmt.numFmt, wb);
  let numFmtId = BUILTIN_NUM_FMT_GENERAL;
  if (code !== null) {
    const id = wb.addNumFmtCode(code);
    if (id < 0) return null;
    numFmtId = id;
  }
  return buildXfRecord(fontIndex, fillIndex, borderIndex, numFmtId, fmt);
}

const RAW_ALIGNMENT_OPTIONAL_FIELDS = [
  'justifyLastLine',
  'hasAlignment',
  'hasHorizontalAlign',
  'hasVerticalAlign',
  'hasWrapText',
  'hasJustifyLastLine',
  'textRotation',
  'indent',
  'relativeIndent',
  'shrinkToFit',
  'readingOrder',
] as const;

const formatHasGroupFields = (fmt: CellFormat, group: CellStyleFormatGroup): boolean =>
  CELL_STYLE_GROUP_FIELDS[group].some((field) => Object.hasOwn(fmt, field));

const FONT_EFFECT_FIELDS = [
  ['fontOutline', 'outline', 'hasOutline'],
  ['fontShadow', 'shadow', 'hasShadow'],
  ['fontCondense', 'condense', 'hasCondense'],
  ['fontExtend', 'extend', 'hasExtend'],
] as const;

const sameFontNonAxisField = (
  format: CellFormat,
  current: Partial<CellFormat>,
  field: keyof CellFormat,
): boolean => {
  // A sparse direct format changes only authored fields. An omitted bold,
  // colour, or vertical-alignment axis inherits the raw source value and
  // therefore must not force a generic font rebuild that drops native
  // metadata. An explicitly authored `false`/`undefined` remains a real
  // reset and is compared below.
  if (!Object.hasOwn(format, field)) return true;
  const left = format[field];
  const right = current[field];
  if (
    field === 'bold' ||
    field === 'italic' ||
    field === 'strike' ||
    FONT_EFFECT_FIELDS.some(([key]) => key === field)
  ) {
    return (left === true) === (right === true);
  }
  return Object.is(left, right);
};

const fontNonAxisMatches = (format: CellFormat, current: Partial<CellFormat>): boolean =>
  CELL_STYLE_GROUP_FIELDS.font.every(
    (field) =>
      field === 'fontFamily' ||
      field === 'fontSize' ||
      sameFontNonAxisField(format, current, field),
  );

/** Preserve the raw inherited font record when a direct write changes only
 *  family and/or size. The raw record carries theme, charset, colour and
 *  apply metadata that CellFormat intentionally does not model. */
const preserveRawFontForAxisEdit = (
  wb: WorkbookHandle,
  record: CellXf,
  source: CellXf,
  currentProjected: Partial<CellFormat>,
  effectiveFormat: CellFormat,
): void => {
  if (!fontNonAxisMatches(effectiveFormat, currentProjected)) return;
  const familyChanged =
    Object.hasOwn(effectiveFormat, 'fontFamily') &&
    effectiveFormat.fontFamily !== currentProjected.fontFamily;
  const sizeChanged =
    Object.hasOwn(effectiveFormat, 'fontSize') &&
    effectiveFormat.fontSize !== currentProjected.fontSize;
  if (!familyChanged && !sizeChanged) return;
  const baseFont = wb.getFontRecord(source.fontIndex);
  if (!baseFont) return;
  const font = { ...baseFont };
  if (familyChanged) {
    const name = effectiveFormat.fontFamily;
    if (name !== undefined) font.name = name;
    font.scheme =
      name === wb.workbookDefaultFont?.name ? wb.workbookDefaultFont?.scheme : undefined;
  }
  if (sizeChanged && effectiveFormat.fontSize !== undefined) {
    font.size = fontCssPxToPoints(effectiveFormat.fontSize);
  }
  if (!font.name || !Number.isFinite(font.size) || font.size <= 0) return;
  const fontIndex = wb.addFontRecord(font);
  if (fontIndex >= 0) {
    record.fontIndex = fontIndex;
    record.applyFont = source.applyFont;
  }
};

const copyRawGroup = (target: CellXf, source: CellXf, group: CellStyleFormatGroup): void => {
  if (group === 'font') {
    target.fontIndex = source.fontIndex;
    target.applyFont = source.applyFont;
  } else if (group === 'fill') {
    target.fillIndex = source.fillIndex;
    target.applyFill = source.applyFill;
  } else if (group === 'border') {
    target.borderIndex = source.borderIndex;
    target.applyBorder = source.applyBorder;
  } else if (group === 'number') {
    target.numFmtId = source.numFmtId;
    target.applyNumberFormat = source.applyNumberFormat;
  } else if (group === 'alignment') {
    target.horizontalAlign = source.horizontalAlign;
    target.verticalAlign = source.verticalAlign;
    target.wrapText = source.wrapText;
    target.applyAlignment = source.applyAlignment;
    const targetRecord = target as unknown as Record<string, unknown>;
    const sourceRecord = source as unknown as Record<string, unknown>;
    for (const field of RAW_ALIGNMENT_OPTIONAL_FIELDS) {
      if (Object.hasOwn(source, field)) {
        targetRecord[field] = sourceRecord[field];
      } else {
        delete targetRecord[field];
      }
    }
  } else if (group === 'protection') {
    target.applyProtection = source.applyProtection;
    target.hasProtection = source.hasProtection;
    target.locked = source.locked;
    target.hidden = source.hidden;
  }
};

const literalColorSpec = (): ColorSpec => ({ kind: 0, rgb: 0, theme: 0, tint: 0, indexed: 0 });

const copyRawAlignmentFieldsExcept = (
  target: CellXf,
  source: CellXf,
  patch: Partial<CellFormat>,
): void => {
  const targetRecord = target as unknown as Record<string, unknown>;
  const sourceRecord = source as unknown as Record<string, unknown>;
  const owners: Readonly<Record<string, keyof CellFormat | undefined>> = {
    horizontalAlign: 'align',
    hasHorizontalAlign: 'align',
    verticalAlign: 'vAlign',
    hasVerticalAlign: 'vAlign',
    wrapText: 'wrap',
    hasWrapText: 'wrap',
    justifyLastLine: 'justifyLastLine',
    hasJustifyLastLine: 'justifyLastLine',
    textRotation: 'rotation',
    indent: 'indent',
    shrinkToFit: 'shrinkToFit',
    readingOrder: 'textDirection',
  };
  for (const [field, owner] of Object.entries(owners)) {
    // `hasAlignment` is derived from all alignment fields. Once any one of
    // them is authored, the target builder owns that flag; copying a source
    // false would otherwise suppress the newly authored field.
    if (Object.hasOwn(patch, owner as keyof CellFormat)) continue;
    if (Object.hasOwn(sourceRecord, field)) targetRecord[field] = sourceRecord[field];
    else delete targetRecord[field];
  }
  const alignmentPatched = [
    'align',
    'vAlign',
    'wrap',
    'justifyLastLine',
    'shrinkToFit',
    'indent',
    'rotation',
    'textDirection',
  ].some((field) => Object.hasOwn(patch, field));
  if (!alignmentPatched) {
    if (Object.hasOwn(sourceRecord, 'hasAlignment')) {
      targetRecord.hasAlignment = sourceRecord.hasAlignment;
    } else {
      delete targetRecord.hasAlignment;
    }
  } else if (targetRecord.hasAlignment === false && sourceRecord.hasAlignment === false) {
    // Keep an explicit empty alignment child when the patch itself clears the
    // last authored alignment field. Otherwise the target builder's `true`
    // presence bit remains authoritative for the newly patched field.
    targetRecord.hasAlignment = false;
  }
  for (const field of RAW_ALIGNMENT_OPTIONAL_FIELDS) {
    if (Object.hasOwn(owners, field)) continue;
    if (field === 'hasAlignment') continue;
    if (Object.hasOwn(sourceRecord, field)) targetRecord[field] = sourceRecord[field];
    else delete targetRecord[field];
  }
};

/**
 * Build a new XF by changing only the fields in `patch` while retaining the
 * raw source record for every untouched style group.  Axis styles are native
 * cell-style XFs rather than CellFormat entries, so routing them through the
 * ordinary sparse writer would otherwise lose apply flags, charset/theme
 * selectors, quotePrefix, or extended alignment metadata.
 */
export function patchXfIndexFromBase(
  wb: WorkbookHandle,
  baseXfIndex: number,
  patch: Partial<CellFormat>,
): number {
  const source = wb.getCellXf(baseXfIndex);
  if (!source) return -1;
  const base = cellFormatFromXf(wb, source, wb.workbookDefaultFont);
  const effective: CellFormat = { ...base, ...patch };
  if (patch.borders) effective.borders = { ...(base.borders ?? {}), ...patch.borders };
  for (const key of Object.keys(patch) as Array<keyof CellFormat>) {
    if (patch[key] === undefined) delete (effective as Partial<CellFormat>)[key];
  }
  withFillReplacement(effective, patch);

  const target = buildXfForFormat(wb, effective);
  if (!target) return -1;
  const has = (key: keyof CellFormat): boolean => Object.hasOwn(patch, key);
  const groups: readonly CellStyleFormatGroup[] = [
    'number',
    'alignment',
    'font',
    'border',
    'fill',
    'protection',
  ];
  const touched = new Set<CellStyleFormatGroup>();
  for (const group of groups) {
    if (CELL_STYLE_GROUP_FIELDS[group].some((key) => has(key))) touched.add(group);
  }
  for (const group of groups) {
    if (!touched.has(group)) copyRawGroup(target, source, group);
  }

  if (touched.has('font')) {
    const sourceFont = wb.getFontRecord(source.fontIndex);
    if (!sourceFont) return -1;
    const font = { ...sourceFont };
    if (has('fontFamily')) {
      const defaultFont = wb.workbookDefaultFont;
      font.name = patch.fontFamily ?? defaultFont?.name ?? font.name;
      font.scheme =
        font.name === wb.workbookDefaultFont?.name ? wb.workbookDefaultFont?.scheme : undefined;
    }
    if (has('fontSize')) {
      const size = patch.fontSize;
      if (size !== undefined && Number.isFinite(size) && size > 0) {
        font.size = fontCssPxToPoints(size);
      } else if (wb.workbookDefaultFont) {
        font.size = wb.workbookDefaultFont.size;
      }
    }
    if (has('bold')) {
      font.bold = patch.bold === true;
      font.hasBold = true;
    }
    if (has('italic')) {
      font.italic = patch.italic === true;
      font.hasItalic = true;
    }
    if (has('strike')) {
      font.strike = patch.strike === true;
      font.hasStrike = true;
    }
    if (has('underline')) {
      font.underline =
        patch.underline === true || patch.underline === 'single'
          ? 1
          : patch.underline === 'double'
            ? 2
            : patch.underline === 'singleAccounting'
              ? 3
              : patch.underline === 'doubleAccounting'
                ? 4
                : 0;
    }
    if (has('fontVertAlign')) {
      font.vertAlign =
        patch.fontVertAlign === 'superscript' ? 1 : patch.fontVertAlign === 'subscript' ? 2 : 0;
    }
    if (has('color')) {
      font.colorArgb =
        patch.color === undefined ? 0xff000000 : (cssColorToArgb(patch.color) ?? 0xff000000);
      // A source theme/indexed selector is authoritative over colorArgb in
      // the native writer. Replacing it with kind=0 makes an explicit UI
      // colour win while leaving every other raw font field intact.
      font.color = literalColorSpec();
    }
    for (const [key, flag, presence] of FONT_EFFECT_FIELDS) {
      if (!has(key)) continue;
      font[flag] = patch[key] === true;
      font[presence] = true;
    }
    const fontIndex = wb.addFontRecord(font);
    if (fontIndex < 0) return -1;
    target.fontIndex = fontIndex;
    target.applyFont = true;
  }
  if (touched.has('fill')) {
    const sourceFill = wb.getFillRecord(source.fillIndex);
    const targetFill = wb.getFillRecord(target.fillIndex);
    if (!sourceFill || !targetFill) return -1;
    // A gradient replaces the whole fill, and a solid or pattern fill
    // replaces a gradient, so either change takes the rebuilt record.
    const replacesFill =
      patch.fillGradient !== undefined ||
      (sourceFill.gradient !== undefined &&
        (has('fill') || has('fillPattern') || has('fillGradient')));
    const fill = replacesFill ? { ...targetFill } : { ...sourceFill };
    const fillRecord = fill as unknown as Record<string, unknown>;
    const setLiteral = (slot: 'fg' | 'bg', argb: number): void => {
      fillRecord[slot] = literalColorSpec();
      fillRecord[`${slot}Argb`] = argb;
    };
    if (!replacesFill) {
      if (has('fillPattern') || has('fill')) fill.pattern = targetFill.pattern;
      if (has('fillPatternColor')) setLiteral('fg', targetFill.fgArgb);
      // `fill` is the visible solid foreground for a solid/empty pattern and
      // the background for a patterned fill. Change only that owned colour
      // pair; the other selector stays exactly as it was in the source record.
      if (has('fill'))
        setLiteral(
          targetFill.pattern > 1 ? 'bg' : 'fg',
          targetFill.pattern > 1 ? targetFill.bgArgb : targetFill.fgArgb,
        );
    }
    const fillIndex = wb.addFillRecord(fill);
    if (fillIndex < 0) return -1;
    target.fillIndex = fillIndex;
    target.applyFill = true;
  }
  if (touched.has('border')) {
    const sourceBorder = wb.getBorderRecord(source.borderIndex);
    const targetBorder = wb.getBorderRecord(target.borderIndex);
    if (!sourceBorder || !targetBorder) return -1;
    const border = { ...sourceBorder };
    const borderPatch = patch.borders ?? {};
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      if (Object.hasOwn(borderPatch, side)) border[side] = targetBorder[side];
    }
    if (Object.hasOwn(borderPatch, 'diagonalDown')) {
      border.diagonalDown = targetBorder.diagonalDown;
      if (borderPatch.diagonalDown) border.diagonal = targetBorder.diagonal;
    }
    if (Object.hasOwn(borderPatch, 'diagonalUp')) {
      border.diagonalUp = targetBorder.diagonalUp;
      if (borderPatch.diagonalUp) border.diagonal = targetBorder.diagonal;
    }
    const borderIndex = wb.addBorderRecord(border);
    if (borderIndex < 0) return -1;
    target.borderIndex = borderIndex;
    target.applyBorder = true;
  }
  if (touched.has('alignment')) copyRawAlignmentFieldsExcept(target, source, patch);
  // These fields are not style groups, but they are meaningful OOXML
  // metadata and must survive an axis edit verbatim.
  const targetRecord = target as unknown as Record<string, unknown>;
  const sourceRecord = source as unknown as Record<string, unknown>;
  for (const key of ['quotePrefix', 'xfId'] as const) {
    if (Object.hasOwn(sourceRecord, key)) targetRecord[key] = sourceRecord[key];
    else delete targetRecord[key];
  }
  return wb.addXfRecord(target);
}

/** Resolve a CellFormat to an engine xfIndex. `styleXfIds` maps a named style
 *  to its `<cellStyleXfs>` row, so a styled cell records which style it came
 *  from instead of collapsing into anonymous direct formatting. Returns -1 on
 *  engine failure. */
function resolveXfForFormat(
  wb: WorkbookHandle,
  fmt: CellFormat,
  styleXfIds: ReadonlyMap<string, number>,
  addr: Addr,
  namedStyleXfs: ReadonlyMap<string, NamedStyleXfResolution>,
  strict: boolean,
  sheetFontPatch?: SheetFontPatch,
): number {
  const effectiveFormat = formatWithSheetFontPatch(fmt, sheetFontPatch);
  const record = buildXfForFormat(wb, fmt, sheetFontPatch);
  if (!record) return -1;
  const namedKey = fmt.cellStyle ?? 'normal';
  const namedXf = namedStyleXfs.get(namedKey);
  const mappedXfId =
    fmt.cellStyle === undefined ? styleXfIds.get('normal') : styleXfIds.get(fmt.cellStyle);
  if (strict && mappedXfId !== undefined && !namedXf) return -1;

  const currentXfIndex = wb.getCellXfIndex(addr.sheet, addr.row, addr.col);
  const currentRaw = currentXfIndex === null ? null : wb.getCellXf(currentXfIndex);
  const currentProjected = currentRaw
    ? cellFormatFromXf(wb, currentRaw, wb.workbookDefaultFont)
    : null;
  if (currentRaw && currentProjected) {
    preserveRawFontForAxisEdit(wb, record, currentRaw, currentProjected, effectiveFormat);
  }
  // quotePrefix has no CellFormat counterpart. Carry it through an edit as
  // opaque XF metadata while the current cell still points at its source row.
  if (currentRaw && Object.hasOwn(currentRaw, 'quotePrefix')) {
    record.quotePrefix = currentRaw.quotePrefix;
  }
  const groups: readonly CellStyleFormatGroup[] = [
    'number',
    'alignment',
    'font',
    'border',
    'fill',
    'protection',
  ];
  for (const group of groups) {
    const namedMatch =
      namedXf?.groups.includes(group) === true &&
      !(group === 'alignment' && rawAlignmentHasAuthoredDifference(effectiveFormat, namedXf.raw)) &&
      cellStyleGroupMatches(effectiveFormat, namedXf.projected, group);
    const currentMatch =
      currentRaw !== null &&
      currentProjected !== null &&
      !(group === 'alignment' && rawAlignmentHasAuthoredDifference(effectiveFormat, currentRaw)) &&
      !(
        group === 'protection' &&
        formatHasGroupFields(fmt, 'protection') &&
        currentRaw.hasProtection !== true
      ) &&
      cellStyleGroupMatches(effectiveFormat, currentProjected, group);
    // An explicitly named style owns the groups it declares, so prefer its
    // raw selectors. An unstyled cell still inherits Normal semantically, but
    // its current XF may carry meaningful source metadata (apply flags,
    // explicit defaults), which must survive an unrelated edit.
    if (fmt.cellStyle !== undefined && namedMatch) copyRawGroup(record, namedXf.raw, group);
    else if (!formatHasGroupFields(effectiveFormat, group) && currentRaw !== null)
      copyRawGroup(record, currentRaw, group);
    else if (currentMatch) copyRawGroup(record, currentRaw, group);
    else if (namedMatch && !formatHasGroupFields(effectiveFormat, group))
      copyRawGroup(record, namedXf.raw, group);
  }
  const xfId = mappedXfId;
  return wb.addXfRecord(xfId === undefined ? record : { ...record, xfId });
}

/** Publish the named-style catalog without touching any cell, for a style-list
 *  change on a store whose cells need no rewrite. */
export function syncNamedStyleCatalogToEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.cellFormatting) return;
  syncNamedCellStylesToEngine(wb, store.getState(), (format) => resolveStyleXf(wb, format));
}

/** Resolve a named style's own formatting to a `<cellStyleXfs>` row. */
function resolveStyleXf(wb: WorkbookHandle, format: CellFormat): number {
  const record = buildXfForFormat(wb, format);
  return record ? wb.addCellStyleXfRecord(record) : -1;
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
