import {
  customPivotTableStyleById,
  pivotTableStyleAssignment,
  tableStyleSwatch,
} from '../commands/format-as-table.js';
import type { CellFormat, SpreadsheetStore } from '../store/store.js';
import { addrKey } from './address.js';
import { syncHyperlinksToEngine } from './format-sync.js';
import {
  BUILTIN_NUM_FMT_GENERAL,
  borderRecordFromFormat,
  borderRecordToFormat,
  buildXfRecord,
  fillRecordFromFormat,
  fillRecordToFormat,
  fontRecordFromFormat,
  fontRecordToFormat,
  formatCodeToNumFmt,
  numFmtToFormatCode,
} from './format-writeback.js';
import type { CellXf } from './types.js';
import { syncValidationsToEngine } from './validation-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

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
): void {
  if (!wb.capabilities.cellFormatting) return;
  const formats = store.getState().format.formats;
  const previous = formattedKeySet(wb, sheet);
  const current = new Set<string>();
  for (const [key, fmt] of formats) {
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    if (Number.parseInt(sStr, 10) !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    const xfIndex = resolveXfForFormat(wb, fmt);
    if (xfIndex < 0) continue;
    wb.setCellXfIndex(sheet, row, col, xfIndex);
    current.add(key);
  }
  // Reset cells that were formatted before but no longer are — the cleared
  // format must not linger in the engine XF table.
  for (const key of previous) {
    if (current.has(key)) continue;
    const [, rStr, cStr] = key.split(':');
    if (rStr === undefined || cStr === undefined) continue;
    wb.setCellXfIndex(sheet, Number.parseInt(rStr, 10), Number.parseInt(cStr, 10), 0);
  }
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
  const updates: Array<{ key: string; patch: Partial<CellFormat> }> = [];
  const physicalCells = wb.physicalCells ? wb.physicalCells(sheet) : wb.cells(sheet);
  for (const c of physicalCells) {
    const xfIndex = wb.getCellXfIndex(sheet, c.addr.row, c.addr.col);
    if (xfIndex === null || xfIndex <= 0) continue;
    const xf = wb.getCellXf(xfIndex);
    if (!xf) continue;
    const patch = cellFormatFromXf(wb, xf);
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

export function cellFormatFromXf(wb: WorkbookHandle, xf: CellXf): Partial<CellFormat> {
  const patch: Partial<CellFormat> = {};
  const font = wb.getFontRecord(xf.fontIndex);
  if (font) Object.assign(patch, fontRecordToFormat(font));
  const fill = wb.getFillRecord(xf.fillIndex);
  if (fill) Object.assign(patch, fillRecordToFormat(fill));
  const border = wb.getBorderRecord(xf.borderIndex);
  if (border) Object.assign(patch, borderRecordToFormat(border));
  if (xf.numFmtId !== BUILTIN_NUM_FMT_GENERAL) {
    const code = wb.getNumFmtCode(xf.numFmtId);
    if (code !== null) {
      const numFmt = formatCodeToNumFmt(code);
      if (numFmt) patch.numFmt = numFmt;
    }
  }
  if (xf.horizontalAlign === 1) patch.align = 'left';
  else if (xf.horizontalAlign === 2) patch.align = 'center';
  else if (xf.horizontalAlign === 3) patch.align = 'right';
  if (xf.verticalAlign === 0) patch.vAlign = 'top';
  else if (xf.verticalAlign === 1) patch.vAlign = 'middle';
  // the desktop default vertical alignment is bottom; do not surface it.
  if (xf.wrapText) patch.wrap = true;
  return patch;
}

/** Resolve a CellFormat to an engine xfIndex by ensuring every component
 *  record exists (dedup-on-add) and assembling the XF. Returns -1 on
 *  engine failure. */
function resolveXfForFormat(wb: WorkbookHandle, fmt: CellFormat): number {
  const fontIndex = wb.addFontRecord(fontRecordFromFormat(fmt));
  if (fontIndex < 0) return -1;
  const fillIndex = wb.addFillRecord(fillRecordFromFormat(fmt));
  if (fillIndex < 0) return -1;
  const borderIndex = wb.addBorderRecord(borderRecordFromFormat(fmt));
  if (borderIndex < 0) return -1;
  const code = numFmtToFormatCode(fmt.numFmt);
  let numFmtId = BUILTIN_NUM_FMT_GENERAL;
  if (code !== null) {
    const id = wb.addNumFmtCode(code);
    if (id < 0) return -1;
    numFmtId = id;
  }
  return wb.addXfRecord(buildXfRecord(fontIndex, fillIndex, borderIndex, numFmtId, fmt));
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
): void {
  syncCellFormatsToEngine(wb, store, sheet);
  syncValidationsToEngine(wb, store, sheet);
  syncHyperlinksToEngine(wb, store, sheet);
}
