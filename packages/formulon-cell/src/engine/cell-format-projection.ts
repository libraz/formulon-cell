import { formatWithPending } from '../store/pending-format.js';
import type { CellFormat } from '../store/store.js';
import { fontPointsToCssPx } from './font-units.js';
import {
  BUILTIN_NUM_FMT_GENERAL,
  borderRecordToFormat,
  fillRecordToFormat,
  fontRecordToFormat,
  formatCodeToNumFmt,
  readingOrderToTextDirection,
  textRotationToCellRotation,
} from './format-writeback.js';
import type { Addr, CellXf, FontRecord } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const HALIGN_GENERAL = 0;
const VALIGN_BOTTOM = 2;

export function cellFormatFromXf(
  wb: WorkbookHandle,
  xf: CellXf,
  workbookDefaultFont: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null = null,
  styleKeysByXfId: ReadonlyMap<number, string> | null = null,
): Partial<CellFormat> {
  return cellFormatFromXfWithStatus(wb, xf, workbookDefaultFont, styleKeysByXfId).format;
}

export function cellFormatFromXfWithStatus(
  wb: WorkbookHandle,
  xf: CellXf,
  workbookDefaultFont: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null,
  styleKeysByXfId: ReadonlyMap<number, string> | null,
): { readonly format: Partial<CellFormat>; readonly complete: boolean } {
  const patch: Partial<CellFormat> = {};
  let complete = true;
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
  else complete = false;
  const fill =
    typeof resolved.getResolvedFillRecord === 'function'
      ? resolved.getResolvedFillRecord(xf.fillIndex)
      : wb.getFillRecord(xf.fillIndex);
  if (fill) Object.assign(patch, fillRecordToFormat(fill));
  else complete = false;
  const border =
    typeof resolved.getResolvedBorderRecord === 'function'
      ? resolved.getResolvedBorderRecord(xf.borderIndex)
      : wb.getBorderRecord(xf.borderIndex);
  if (border) Object.assign(patch, borderRecordToFormat(border));
  else complete = false;
  if (xf.numFmtId !== BUILTIN_NUM_FMT_GENERAL) {
    const code = wb.getNumFmtCode(xf.numFmtId);
    if (code === null) complete = false;
    else {
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
  return { format: patch, complete };
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
  const projection = projectDisplayCellFormatFromXf(wb, effective.xfIndex);
  if (!projection) return null;
  const format = projection.format;
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

/** Project only the display format carried by one native cell XF. Row/column
 *  ranges expose XF indices without cell addresses, so this deliberately
 *  omits effective-style protection and number-format completion. */
export function displayCellFormatFromXf(
  wb: WorkbookHandle,
  xfIndex: number,
): Partial<CellFormat> | null {
  const projection = projectDisplayCellFormatFromXf(wb, xfIndex);
  return projection?.complete ? projection.format : null;
}

function projectDisplayCellFormatFromXf(
  wb: WorkbookHandle,
  xfIndex: number,
): { readonly format: Partial<CellFormat>; readonly complete: boolean } | null {
  const xf = wb.getCellXf(xfIndex);
  if (!xf) return null;
  const projection = cellFormatFromXfWithStatus(wb, xf, wb.workbookDefaultFont, null);
  const format = projection.format;
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
  return { format, complete: projection.complete };
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
  cache?: Map<string, Partial<CellFormat> | null>,
): Partial<CellFormat> | null {
  const effective = effectiveCellFormatFromEngine(wb, addr.sheet, addr.row, addr.col, cache);
  const direct = formatWithPending(state, addr);
  if (!effective && !direct) return null;
  return { ...(effective ?? {}), ...(direct ?? {}) };
}
