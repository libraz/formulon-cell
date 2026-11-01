import { parseAddrKey } from '../engine/address.js';
import type { CellFormat } from '../store/store.js';

/** The font fields that decide a run's descent. */
export type RowBaselineFont = Pick<CellFormat, 'fontFamily' | 'fontSize' | 'bold' | 'italic'>;

const bySheet = new WeakMap<
  ReadonlyMap<string, CellFormat>,
  Map<number, Map<number, RowBaselineFont>>
>();

const bottomAligned = (format: CellFormat): boolean =>
  format.vAlign === undefined || format.vAlign === 'bottom';

/**
 * Per row, the largest explicit font among bottom-aligned cells — empty
 * formatted cells and cells off screen included. The spreadsheet seats every
 * bottom-aligned run in a row on one baseline, raised by the deepest descent
 * among them; top- and middle-aligned cells neither join nor move it.
 */
export function rowBaselineFonts(
  formats: ReadonlyMap<string, CellFormat>,
  sheet: number,
): ReadonlyMap<number, RowBaselineFont> {
  let sheets = bySheet.get(formats);
  if (!sheets) {
    sheets = new Map();
    bySheet.set(formats, sheets);
  }
  const cached = sheets.get(sheet);
  if (cached) return cached;
  const rows = new Map<number, RowBaselineFont>();
  for (const [key, format] of formats) {
    const size = format.fontSize;
    if (size === undefined || !Number.isFinite(size) || !bottomAligned(format)) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const current = rows.get(addr.row);
    if (current?.fontSize !== undefined && current.fontSize >= size) continue;
    rows.set(addr.row, {
      fontSize: size,
      fontFamily: format.fontFamily,
      bold: format.bold,
      italic: format.italic,
    });
  }
  sheets.set(sheet, rows);
  return rows;
}
