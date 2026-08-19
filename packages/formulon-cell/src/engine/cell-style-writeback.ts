import { CELL_STYLES, type CellStyleId, getCellStyle } from '../commands/cell-styles.js';
import type { CellFormat, State } from '../store/store.js';
import type { WorkbookHandle } from './workbook-handle.js';

/**
 * OOXML `<cellStyle>` metadata for the gallery's built-in styles. The gallery
 * labels already match Excel's style names everywhere except "Warning", which
 * OOXML spells "Warning Text".
 *
 * `builtinId` is the ordinal from the OOXML built-in cell-style table. The
 * engine validates it as 0..47, so "Explanatory Text" (53) has no usable
 * ordinal and is written as a named style without a built-in counterpart.
 */
const BUILTIN_STYLE_IDS: Partial<Record<CellStyleId, number>> = {
  normal: 0,
  comma: 3,
  currency: 4,
  percent: 5,
  comma0: 6,
  currency0: 7,
  note: 10,
  warning: 11,
  title: 15,
  heading1: 16,
  heading2: 17,
  heading3: 18,
  heading4: 19,
  inputCell: 20,
  outputCell: 21,
  calculation: 22,
  checkCell: 23,
  linkedCell: 24,
  totalCell: 25,
  good: 26,
  bad: 27,
  neutral: 28,
  accent1: 29,
  accent1_20: 30,
  accent2: 33,
  accent2_20: 34,
  accent3: 37,
  accent3_20: 38,
  accent4: 41,
  accent4_20: 42,
  accent5: 45,
  accent5_20: 46,
  accent6: 49,
  accent6_20: 50,
};

const OOXML_STYLE_NAMES: Partial<Record<CellStyleId, string>> = {
  warning: 'Warning Text',
};

/** The `<cellStyleXfs>` row every unstyled cell inherits from. OOXML reserves
 *  index 0 for it, and a cell XF that names no style leaves `xfId` at 0. */
const NORMAL_STYLE_XF_ID = 0;

/** A named style as it is about to be written: the store-side key that cells
 *  carry in `CellFormat.cellStyle`, the OOXML name, and the format the style
 *  itself defines. */
export interface NamedStyleRegistration {
  key: string;
  name: string;
  builtinId: number | null;
  format: Partial<CellFormat>;
}

const ooxmlNameFor = (id: CellStyleId, label: string): string => OOXML_STYLE_NAMES[id] ?? label;

/** Reverse of `ooxmlNameFor` + the gallery labels: maps an OOXML style name
 *  back onto the gallery id a hydrated cell should carry, so a style survives
 *  a save/load without turning into a look-alike custom entry. */
export function cellStyleKeyForOoxmlName(name: string): string {
  const needle = name.trim().toLowerCase();
  for (const def of CELL_STYLES) {
    if (ooxmlNameFor(def.id, def.label).toLowerCase() === needle) return def.id;
  }
  return name;
}

/** Every named style referenced by at least one cell on the sheet set, in a
 *  stable order. Styles nothing references are not written — Excel keeps the
 *  full built-in gallery in every file, but emitting 35 unused entries would
 *  bloat a workbook the user never styled. */
export function collectNamedStyles(state: State): NamedStyleRegistration[] {
  const keys = new Set<string>();
  for (const fmt of state.format.formats.values()) {
    if (fmt.cellStyle) keys.add(fmt.cellStyle);
  }
  const out: NamedStyleRegistration[] = [];
  for (const key of keys) {
    const builtin = getCellStyle(key as CellStyleId);
    if (builtin) {
      out.push({
        key,
        name: ooxmlNameFor(builtin.id, builtin.label),
        builtinId: BUILTIN_STYLE_IDS[builtin.id] ?? null,
        format: builtin.format,
      });
      continue;
    }
    const custom = state.format.customCellStyles?.find((s) => s.label === key);
    if (custom) out.push({ key, name: custom.label, builtinId: null, format: custom.format });
  }
  return out;
}

/**
 * Register every named style the sheet uses as a real `<cellStyle>` +
 * `<cellStyleXfs>` pair and return the style-xf index each store-side key
 * resolved to. Callers hand that map to the cell XF writeback, which points
 * each styled cell's `xfId` at its style — the two-layer model Excel needs for
 * "edit the style, every cell using it follows".
 *
 * Returns an empty map when the engine cannot author named styles; cells then
 * keep their direct formatting and simply inherit the default style.
 */
export function syncNamedCellStylesToEngine(
  wb: WorkbookHandle,
  state: State,
  resolveStyleXf: (format: Partial<CellFormat>) => number,
): Map<string, number> {
  const resolved = new Map<string, number>();
  if (!wb.capabilities.cellStyleMutate) return resolved;
  const registrations = collectNamedStyles(state);
  if (registrations.length === 0) return resolved;

  // A workbook that has never carried a named style has an empty
  // `<cellStyleXfs>` table, so the first row added would land at index 0 —
  // the row every unstyled cell inherits. Seed Normal there first.
  if (wb.cellStyleXfCount() === 0) {
    const normalXf = resolveStyleXf({});
    if (normalXf !== NORMAL_STYLE_XF_ID) return resolved;
    wb.setNamedCellStyle('Normal', NORMAL_STYLE_XF_ID, BUILTIN_STYLE_IDS.normal ?? null);
  }

  for (const reg of registrations) {
    const xfId = resolveStyleXf(reg.format);
    if (xfId < 0) continue;
    if (!wb.setNamedCellStyle(reg.name, xfId, reg.builtinId)) continue;
    resolved.set(reg.key, xfId);
  }
  return resolved;
}

/** Map each `<cellStyleXfs>` row back to the gallery key a hydrated cell
 *  should carry. Built-in ordinal 0 ("Normal") is left out: an unstyled cell
 *  points at it and must not come back tagged with a style. */
export function cellStyleKeysByXfId(wb: WorkbookHandle): Map<number, string> {
  const out = new Map<number, string>();
  if (!wb.capabilities.cellStyles) return out;
  for (const style of wb.getNamedCellStyles()) {
    if (style.xfId === NORMAL_STYLE_XF_ID) continue;
    if (!out.has(style.xfId)) out.set(style.xfId, cellStyleKeyForOoxmlName(style.name));
  }
  return out;
}
