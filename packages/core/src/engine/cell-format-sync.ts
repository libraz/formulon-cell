import type { CellFormat, SpreadsheetStore } from '../store/store.js';
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
import { syncValidationsToEngine } from './validation-sync.js';
import { addrKey } from './workbook-handle.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * Push every format entry on `sheet` from FormatSlice into the engine's XF
 * table. For each cell the writeback ensures a font / fill / border / numFmt
 * record exists (the engine dedups against existing rows), assembles an XF,
 * and pins the cell's `xfIndex`. No-op when `capabilities.cellFormatting`
 * is off.
 *
 * The sync is "FormatSlice → engine"; entries that have been deleted from the
 * store reset to xfIndex 0 (the workbook default). Stub engines bypass the
 * whole walk because every wrapper short-circuits on the capability flag.
 */
export function syncCellFormatsToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.cellFormatting) return;
  const formats = store.getState().format.formats;
  // Track which cells we wrote so we can clear stale XF assignments on cells
  // whose entry was removed since the last sync. We do not currently track
  // a high-water mark of previously-formatted cells, so on first sync after
  // a clear the store-side delete already drove the entry away — the engine
  // keeps its old xfIndex unless the user re-formats. This matches Excel:
  // "Clear Formats" sets xf to 0; that path lives in the future.
  for (const [key, fmt] of formats) {
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    if (Number.parseInt(sStr, 10) !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    const xfIndex = resolveXfForFormat(wb, fmt);
    if (xfIndex < 0) continue;
    wb.setCellXfIndex(sheet, row, col, xfIndex);
  }
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
  for (const c of wb.cells(sheet)) {
    const xfIndex = wb.getCellXfIndex(sheet, c.addr.row, c.addr.col);
    if (xfIndex === null || xfIndex <= 0) continue;
    const xf = wb.getCellXf(xfIndex);
    if (!xf) continue;
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
    // Excel default vertical alignment is bottom; do not surface it.
    if (xf.wrapText) patch.wrap = true;
    if (Object.keys(patch).length > 0) {
      updates.push({ key: addrKey(c.addr), patch });
    }
  }
  if (updates.length === 0) return;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const u of updates) {
      const prev = formats.get(u.key) ?? {};
      formats.set(u.key, { ...prev, ...u.patch });
    }
    return { ...s, format: { formats } };
  });
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
