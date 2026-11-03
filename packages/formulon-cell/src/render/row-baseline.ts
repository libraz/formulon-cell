import { parseAddrKey } from '../engine/address.js';
import type { CellFormat } from '../store/store.js';

/** The font fields that decide a run's descent, and whether the row holds an
 *  accounting underline, which raises the shared baseline further. */
export type RowBaselineFont = Pick<CellFormat, 'fontFamily' | 'fontSize' | 'bold' | 'italic'> & {
  accountingUnderline?: boolean;
};

const bySheet = new WeakMap<
  ReadonlyMap<string, CellFormat>,
  Map<number, Map<number, RowBaselineFont>>
>();

const bottomAligned = (format: CellFormat): boolean =>
  format.vAlign === undefined || format.vAlign === 'bottom';

const accountingUnderline = (format: CellFormat): boolean =>
  format.underline === 'singleAccounting' || format.underline === 'doubleAccounting';

/**
 * Per row, the largest explicit font among bottom-aligned cells — empty
 * formatted cells and cells off screen included — and whether any of them
 * carries an accounting underline. The spreadsheet seats every bottom-aligned
 * run in a row on one baseline, raised by the deepest descent among them and
 * by the accounting underline's band; top- and middle-aligned cells neither
 * join nor move it.
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
    if (!bottomAligned(format)) continue;
    const size = format.fontSize;
    const sized = size !== undefined && Number.isFinite(size);
    const accounting = accountingUnderline(format);
    if (!sized && !accounting) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const current = rows.get(addr.row);
    const larger = sized && (current?.fontSize === undefined || current.fontSize < size);
    const font: RowBaselineFont = larger
      ? { fontSize: size, fontFamily: format.fontFamily, bold: format.bold, italic: format.italic }
      : { ...current };
    if (accounting || current?.accountingUnderline) font.accountingUnderline = true;
    rows.set(addr.row, font);
  }
  sheets.set(sheet, rows);
  return rows;
}
