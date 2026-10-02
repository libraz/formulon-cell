import { type CellStyleFormatGroup, cellStyleGroupMatches } from '../commands/cell-styles.js';
import {
  customPivotTableStyleById,
  pivotTableStyleAssignment,
  tableStyleSwatch,
} from '../commands/format-as-table.js';
import type { CellFormat, SpreadsheetStore } from '../store/store.js';
import { addrKey } from './address.js';
import {
  cellStyleKeysByXfId,
  type NamedStyleXfResolution,
  planNamedCellStylesToEngine,
  resolveNamedStyleXfs,
  syncNamedCellStylesToEngine,
} from './cell-style-writeback.js';
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
  readingOrderToTextDirection,
  textRotationToCellRotation,
} from './format-writeback.js';
import type { Addr, CellValue, CellXf, FontRecord, PhoneticRun } from './types.js';
import { syncValidationsToEngine } from './validation-sync.js';
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
  for (const [key, fmt] of formats) {
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    if (Number.parseInt(sStr, 10) !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    const xfIndex = resolveXfForFormat(
      wb,
      fmt,
      styleXfIds,
      { sheet, row, col },
      namedStyleXfs,
      strict,
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
    current.add(key);
  }
  // Reset cells that were formatted before but no longer are — the cleared
  // format must not linger in the engine XF table.
  for (const key of previous) {
    if (current.has(key)) continue;
    const [, rStr, cStr] = key.split(':');
    if (rStr === undefined || cStr === undefined) continue;
    const resetSucceeded = wb.setCellXfIndex(
      sheet,
      Number.parseInt(rStr, 10),
      Number.parseInt(cStr, 10),
      0,
    );
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
  const updates: Array<{ key: string; patch: Partial<CellFormat> }> = [];
  const physicalCells = wb.physicalCells ? wb.physicalCells(sheet) : wb.cells(sheet);
  for (const c of physicalCells) {
    const phonetic = readPhoneticRuns(wb, c.addr, c.value);
    const xfIndex = wb.getCellXfIndex(sheet, c.addr.row, c.addr.col);
    if (xfIndex === null || xfIndex <= 0) {
      if (phonetic) updates.push({ key: addrKey(c.addr), patch: { phonetic } });
      continue;
    }
    const xf = wb.getCellXf(xfIndex);
    if (!xf) continue;
    const patch = cellFormatFromXf(wb, xf, workbookDefaultFont, styleKeys);
    if (phonetic) patch.phonetic = phonetic;
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
  const font = wb.getFontRecord(xf.fontIndex);
  if (font) Object.assign(patch, fontRecordToFormat(font, workbookDefaultFont));
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
  return patch;
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
function buildXfForFormat(wb: WorkbookHandle, fmt: CellFormat): CellXf | null {
  const fontIndex = wb.addFontRecord(fontRecordFromFormat(fmt, wb.workbookDefaultFont));
  if (fontIndex < 0) return null;
  const fillIndex = wb.addFillRecord(fillRecordFromFormat(fmt));
  if (fillIndex < 0) return null;
  const borderIndex = wb.addBorderRecord(borderRecordFromFormat(fmt));
  if (borderIndex < 0) return null;
  const code = numFmtToFormatCode(fmt.numFmt);
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

const copyRawGroup = (target: CellXf, source: CellXf, group: CellStyleFormatGroup): void => {
  if (group === 'font') target.fontIndex = source.fontIndex;
  else if (group === 'fill') target.fillIndex = source.fillIndex;
  else if (group === 'border') target.borderIndex = source.borderIndex;
  else if (group === 'number') target.numFmtId = source.numFmtId;
  else if (group === 'alignment') {
    target.horizontalAlign = source.horizontalAlign;
    target.verticalAlign = source.verticalAlign;
    target.wrapText = source.wrapText;
    const targetRecord = target as unknown as Record<string, unknown>;
    const sourceRecord = source as unknown as Record<string, unknown>;
    for (const field of RAW_ALIGNMENT_OPTIONAL_FIELDS) {
      if (Object.hasOwn(source, field)) {
        targetRecord[field] = sourceRecord[field];
      } else {
        delete targetRecord[field];
      }
    }
  }
};

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
): number {
  const record = buildXfForFormat(wb, fmt);
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
      !(group === 'alignment' && rawAlignmentHasAuthoredDifference(fmt, namedXf.raw)) &&
      cellStyleGroupMatches(fmt, namedXf.projected, group);
    const currentMatch =
      currentRaw !== null &&
      currentProjected !== null &&
      !(group === 'alignment' && rawAlignmentHasAuthoredDifference(fmt, currentRaw)) &&
      cellStyleGroupMatches(fmt, currentProjected, group);
    if (namedMatch) copyRawGroup(record, namedXf.raw, group);
    else if (currentMatch) copyRawGroup(record, currentRaw, group);
  }
  const xfId = mappedXfId;
  return wb.addXfRecord(xfId === undefined ? record : { ...record, xfId });
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
): void {
  syncCellFormatsToEngine(wb, store, sheet, options);
  syncValidationsToEngine(wb, store, sheet, options);
  syncHyperlinksToEngine(wb, store, sheet, options);
}
