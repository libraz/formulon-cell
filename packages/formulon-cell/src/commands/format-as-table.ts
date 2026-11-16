import type { StyleElementFormats } from '../engine/table-style-elements.js';
import { defaultTotals } from '../engine/table-totals-defaults.js';
import type { CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeArea, rangesIntersect } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { inputLocaleForWorkbook } from './coerce-input.js';
import type { History } from './history.js';
import {
  blockedByProtection,
  isCellWritable,
  TABLE_FIELDS_ALLOWED_ON_PROTECTED_SHEET,
} from './protection.js';
import { recordTablesChange } from './slice-history.js';

/** "Format As Table" overlay. With a table-authoring engine every overlay is
 * backed by an engine table of the same `name`, which the store mirror keeps
 * in step; without one it decorates a plain range for the session only. */
export type TableStyle = 'light' | 'medium' | 'dark';

/** The elements a custom table style paints: its family and hue. */
export function approximateTableStyleElements(
  style: TableStyle,
  color?: string,
): StyleElementFormats {
  const swatch = tableStyleSwatch(style, color);
  const header = { fill: swatch.header, color: swatch.headerText, bold: true };
  return {
    headerRow: header,
    totalRow: header,
    firstRowStripe: { fill: swatch.band },
    firstColumnStripe: { fill: swatch.band },
    firstColumn: { bold: true },
    lastColumn: { bold: true },
  };
}

/** The elements a gallery pick (family, hue, variant) is authored as: whole-table border in
 *  the hue, header fill and text, bold totals, optional stripe and first-column elements. */
export function galleryStyleElements(
  style: TableStyle,
  color: string,
  variant: TableVariant,
): StyleElementFormats {
  const swatch = tableStyleSwatch(style, color);
  const { banded, firstCol } = tableVariantOptions(variant);
  const side = { style: 'thin' as const, color };
  return {
    wholeTable: { borders: { top: side, bottom: side, left: side, right: side } },
    headerRow: { fill: swatch.header, color: swatch.headerText, bold: true },
    totalRow: { bold: true },
    ...(banded ? { firstRowStripe: { fill: swatch.band } } : {}),
    ...(firstCol ? { firstColumn: { bold: true } } : {}),
  };
}

/** Built-in table-style hues — one column per swatch in the style gallery,
 *  ordered to mirror the desktop spreadsheet's themed accent columns. */
const TABLE_STYLE_COLORS: readonly string[] = [
  '#808080',
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
];

/** Fallback hue for overlays created before the gallery existed, or loaded
 *  from the engine without an explicit color. */
export const DEFAULT_TABLE_COLOR = '#4472c4';

const parseHex = (hex: string): [number, number, number] => {
  const h = hex.replace('#', '');
  return [
    Number.parseInt(h.slice(0, 2), 16),
    Number.parseInt(h.slice(2, 4), 16),
    Number.parseInt(h.slice(4, 6), 16),
  ];
};
const toHexByte = (n: number): string =>
  Math.max(0, Math.min(255, Math.round(n)))
    .toString(16)
    .padStart(2, '0');
const mixHex = (hex: string, toward: string, t: number): string => {
  const [r, g, b] = parseHex(hex);
  const [tr, tg, tb] = parseHex(toward);
  return `#${toHexByte(r + (tr - r) * t)}${toHexByte(g + (tg - g) * t)}${toHexByte(b + (tb - b) * t)}`;
};

/** Resolved fills for a table style — shared by the grid painter and the
 *  style-gallery thumbnails so the preview always matches the applied look. */
export interface TableStyleSwatch {
  base: string;
  header: string;
  headerText: string;
  band: string;
}

/** Banded-rows / first-column option pair of a table, as one gallery value. */
export type TableVariant = 'plain' | 'banded' | 'firstCol' | 'bandedFirstCol';

/** The stripe elements whose band size a custom style can set. */
export type StripeElement =
  | 'firstRowStripe'
  | 'secondRowStripe'
  | 'firstColumnStripe'
  | 'secondColumnStripe';

/** A custom table or PivotTable style: a format per element, plus stripe band sizes. */
export interface CustomTableStyle {
  id: string;
  label: string;
  elements: StyleElementFormats;
  /** Rows or columns per stripe band; absent or 1 is one row or column. */
  stripeSizes?: Partial<Record<StripeElement, number>>;
}

export interface PivotTableStyleAssignment {
  sheetIndex: number;
  pivotIndex: number;
  styleId: string;
}

export interface CreateCustomTableStyleOptions {
  /** The style's elements; derived from the active table's look when omitted. */
  elements?: StyleElementFormats;
  stripeSizes?: CustomTableStyle['stripeSizes'];
}

/** Derive the header / banded-row fills for a `(style, color)` pair. */
function tableStyleSwatch(
  style: TableStyle,
  color: string = DEFAULT_TABLE_COLOR,
): TableStyleSwatch {
  if (style === 'light') {
    return {
      base: color,
      header: mixHex(color, '#ffffff', 0.72),
      headerText: '#1f1f1f',
      band: mixHex(color, '#ffffff', 0.92),
    };
  }
  if (style === 'dark') {
    return {
      base: color,
      header: mixHex(color, '#000000', 0.45),
      headerText: '#ffffff',
      band: mixHex(color, '#ffffff', 0.8),
    };
  }
  return {
    base: color,
    header: color,
    headerText: '#ffffff',
    band: mixHex(color, '#ffffff', 0.84),
  };
}

/** Totals-row function of a table column; the engine's `ST_TotalsRowFunction`. */
export type TableTotalsFunction =
  | 'none'
  | 'sum'
  | 'min'
  | 'max'
  | 'average'
  | 'count'
  | 'countNums'
  | 'stdDev'
  | 'var'
  | 'custom';

/** Totals-row settings of one table column. `formula` (no leading `=`) is the
 *  `custom` function's formula; `label` shows when the function is `none`. */
export interface TableColumnTotals {
  function: TableTotalsFunction;
  label?: string;
  formula?: string;
}

export interface TableOverlay {
  /** Stable id used by mutators / pointer routing. */
  id: string;
  /** Source of the overlay. Engine overlays are backed by the engine table
   *  `name`; session overlays exist only in the UI (no table-authoring engine). */
  source: 'engine' | 'session';
  /** Engine table name backing an engine overlay. */
  name?: string;
  /** Range covered by the table including the header row and (optionally)
   *  the total row. */
  range: Range;
  style: TableStyle;
  /** Hue for the style. Optional — engine-loaded and pre-gallery overlays
   *  fall back to `DEFAULT_TABLE_COLOR`. */
  color?: string;
  /** Render the first row as a header (bold + tinted). Defaults to true. */
  showHeader: boolean;
  /** Render the last row as a total row (bold + tinted). Defaults to false. */
  showTotal: boolean;
  /** Apply zebra fills to data rows. Defaults to true. */
  banded: boolean;
  /** Emphasize the first column with bold text. Defaults to false. */
  firstCol?: boolean;
  /** Emphasize the last column with bold text. Defaults to false. */
  lastCol?: boolean;
  /** Column stripes of the engine table; carried so a loaded table keeps it. */
  colStripes?: boolean;
  /** Show the header row's filter buttons (the table's AutoFilter). Absent
   *  means on, as a new table has them. */
  filterButton?: boolean;
  /** Totals-row settings by column, left to right; engine overlays only. */
  totals?: readonly TableColumnTotals[];
  /** Exact engine style name (built-in or custom) the look was read from or
   *  written as. Absent means the name derived from `style` and `color`; an
   *  empty string means the table has no style. */
  styleName?: string;
}

export interface FormatAsTableOptions {
  id?: string;
  style?: TableStyle;
  color?: string;
  showHeader?: boolean;
  showTotal?: boolean;
  banded?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  styleName?: string;
  /** Name of a new engine-backed table, already checked with `tableNameError`;
   *  defaults to the next free `Table<n>`. */
  name?: string;
  /** Label a new total row carries in its first column; defaults to `Total`. */
  totalsLabel?: string;
  /** When the active engine supports it, back the overlay with an OOXML
   * table. The store mirror of the workbook's attached store creates it. */
  workbook?: WorkbookHandle;
}

export interface TableHeaderInferenceWorkbook {
  getValue(addr: { sheet: number; row: number; col: number }): CellValue;
}

const CUSTOM_TABLE_STYLE_PREFIX = 'custom-table:';
const CUSTOM_PIVOT_TABLE_STYLE_PREFIX = 'custom-pivot-table:';
const MAX_TABLE_HEADER_INFERENCE_CELLS = 100_000;

const isNonEmptyTextValue = (value: CellValue): boolean =>
  value.kind === 'text' && value.value.trim().length > 0;

const isNonBlankValue = (value: CellValue): boolean => value.kind !== 'blank';

const BUILTIN_STYLE_PREFIX: Record<TableStyle, string> = {
  light: 'TableStyleLight',
  medium: 'TableStyleMedium',
  dark: 'TableStyleDark',
};
const BUILTIN_STYLE_COUNT: Record<TableStyle, number> = { light: 21, medium: 28, dark: 11 };

/** Index into `TABLE_STYLE_COLORS` nearest to `color`: no color, then the six accents. */
const paletteIndex = (color: string): number => {
  const [r, g, b] = parseHex(color);
  let best = 0;
  let bestDistance = Number.POSITIVE_INFINITY;
  TABLE_STYLE_COLORS.forEach((swatch, i) => {
    const [pr, pg, pb] = parseHex(swatch);
    const distance = (r - pr) ** 2 + (g - pg) ** 2 + (b - pb) ** 2;
    if (distance < bestDistance) {
      best = i;
      bestDistance = distance;
    }
  });
  return best;
};

/** One built-in table style: its family, 1-based number and engine name. */
export interface BuiltinTableStyle {
  family: TableStyle;
  index: number;
  name: string;
}

/** Every built-in table style, in gallery order: Light 1-21, Medium 1-28, Dark 1-11. */
export const BUILTIN_TABLE_STYLES: readonly BuiltinTableStyle[] = (
  ['light', 'medium', 'dark'] as const
).flatMap((family) =>
  Array.from({ length: BUILTIN_STYLE_COUNT[family] }, (_, i) => ({
    family,
    index: i + 1,
    name: `${BUILTIN_STYLE_PREFIX[family]}${i + 1}`,
  })),
);

/** A section of the table style gallery. */
export type TableStyleGallerySection<C = CustomTableStyle> =
  | { readonly kind: 'custom'; readonly styles: readonly C[] }
  | { readonly kind: TableStyle; readonly styles: readonly BuiltinTableStyle[] };

/** Which gallery a section list is for: the Table Design and PivotTable Design galleries start
 *  their Light section with None; Home > Format as Table does not. */
export interface StyleGallerySurface {
  readonly none: boolean;
}

/** The None entry of the table style gallery: no style. */
const TABLE_STYLE_NONE: BuiltinTableStyle = { family: 'light', index: 0, name: '' };

/** The table style gallery's sections in order, shared by the menus and the in-ribbon strip:
 *  Custom first, when there is one, then Light (after None on `surface.none`), Medium and Dark. */
export function tableStyleGallerySections<C>(
  customs: readonly C[],
  surface: StyleGallerySurface,
): TableStyleGallerySection<C>[] {
  return [
    ...(customs.length > 0 ? [{ kind: 'custom' as const, styles: customs }] : []),
    ...(['light', 'medium', 'dark'] as const).map((kind) => ({
      kind,
      styles: [
        ...(kind === 'light' && surface.none ? [TABLE_STYLE_NONE] : []),
        ...BUILTIN_TABLE_STYLES.filter((entry) => entry.family === kind),
      ],
    })),
  ];
}

/** A section of the PivotTable style gallery; an empty name is None. */
export type PivotStyleGallerySection<C = CustomTableStyle> =
  | { readonly kind: 'custom'; readonly styles: readonly C[] }
  | { readonly kind: TableStyle; readonly names: readonly string[] };

const PIVOT_FAMILY_NAMES = { light: 'Light', medium: 'Medium', dark: 'Dark' } as const;

/** The PivotTable style gallery's sections in order, shared by its strip and menu: Custom first,
 *  as in the table gallery, then Light 1-28 (after None on `surface.none`), Medium 1-28 and
 *  Dark 1-28. */
export function pivotStyleGallerySections<C>(
  customs: readonly C[],
  surface: StyleGallerySurface,
): PivotStyleGallerySection<C>[] {
  return [
    ...(customs.length > 0 ? [{ kind: 'custom' as const, styles: customs }] : []),
    ...(['light', 'medium', 'dark'] as const).map((kind) => ({
      kind,
      names: [
        ...(kind === 'light' && surface.none ? [''] : []),
        ...Array.from({ length: 28 }, (_, i) => `PivotStyle${PIVOT_FAMILY_NAMES[kind]}${i + 1}`),
      ],
    })),
  ];
}

/** Built-in engine style name for a gallery family and hue: the first block
 *  of each family, one style per palette column. */
export const tableStyleName = (style: TableStyle, color: string = DEFAULT_TABLE_COLOR): string =>
  `${BUILTIN_STYLE_PREFIX[style]}${paletteIndex(color) + 1}`;

/** Gallery family and hue of a built-in engine style name; null for any
 *  other name. Every block of a family maps onto its palette columns. */
export const parseTableStyleName = (name: string): { style: TableStyle; color: string } | null => {
  for (const style of ['light', 'medium', 'dark'] as const) {
    const prefix = BUILTIN_STYLE_PREFIX[style];
    if (!name.startsWith(prefix)) continue;
    const digits = name.slice(prefix.length);
    if (!/^[1-9]\d*$/.test(digits)) return null;
    const n = Number(digits);
    if (n > BUILTIN_STYLE_COUNT[style]) return null;
    // Dark8-11 follow the neutral / accent1+2 / accent3+4 / accent5+6 pairs; take the first of each.
    const column = style === 'dark' && n > 7 ? ([0, 1, 3, 5][n - 8] ?? 0) : (n - 1) % 7;
    return { style, color: TABLE_STYLE_COLORS[column] ?? DEFAULT_TABLE_COLOR };
  }
  return null;
};

/** Engine style name an overlay is written as. */
export const overlayStyleName = (
  table: Pick<TableOverlay, 'style' | 'color' | 'styleName'>,
): string => table.styleName ?? tableStyleName(table.style, table.color);

interface DefaultObjectNames {
  table: string;
  column: string;
  pivot: string;
}

const ENGLISH_NAMES: DefaultObjectNames = { table: 'Table', column: 'Column', pivot: 'PivotTable' };

/** Names a new table, its blank headers and a new PivotTable start with, as the
 *  desktop app in the workbook's language spells them. PivotTable names are
 *  measured for ja only, and the zh-TW table name stays English unmeasured. */
const DEFAULT_NAMES: Record<string, DefaultObjectNames> = {
  'ja-JP': { table: 'テーブル', column: '列', pivot: 'ピボットテーブル' },
  'de-DE': { ...ENGLISH_NAMES, table: 'Tabelle', column: 'Spalte' },
  'fr-FR': { ...ENGLISH_NAMES, table: 'Tableau', column: 'Colonne' },
  'zh-CN': { ...ENGLISH_NAMES, table: '表', column: '列' },
  'zh-TW': { ...ENGLISH_NAMES, column: '欄' },
  'ko-KR': { ...ENGLISH_NAMES, table: '표', column: '열' },
  'th-TH': { ...ENGLISH_NAMES, column: 'คอลัมน์' },
};

export const defaultObjectNames = (wb: WorkbookHandle): DefaultObjectNames =>
  DEFAULT_NAMES[inputLocaleForWorkbook(wb) ?? ''] ?? ENGLISH_NAMES;

/** Unique column names for a table over `range`, from its header row when present. */
export const tableColumns = (
  workbook: Pick<WorkbookHandle, 'getValue'>,
  range: Range,
  hasHeaders: boolean,
  columnLabel: string,
): string[] => {
  const used = new Set<string>();
  const columns: string[] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    const value = hasHeaders
      ? workbook.getValue({ sheet: range.sheet, row: range.r0, col })
      : { kind: 'blank' as const };
    const base =
      value.kind === 'text' && value.value.trim()
        ? value.value.trim()
        : `${columnLabel}${col - range.c0 + 1}`;
    let name = base;
    let suffix = 2;
    while (used.has(name.toLocaleLowerCase())) name = `${base}${suffix++}`;
    used.add(name.toLocaleLowerCase());
    columns.push(name);
  }
  return columns;
};

/** The next free default table name (`Table1`, `テーブル1`, …) in the workbook's language. */
export const nextTableName = (
  workbook: WorkbookHandle,
  overlays: readonly TableOverlay[],
): string => {
  const existing = new Set(workbook.getTables().map((table) => table.name.toLocaleLowerCase()));
  for (const overlay of overlays) if (overlay.name) existing.add(overlay.name.toLocaleLowerCase());
  const { table } = defaultObjectNames(workbook);
  let n = 1;
  while (existing.has(`${table}${n}`.toLocaleLowerCase())) n += 1;
  return `${table}${n}`;
};

/** Excel pre-checks "My table has headers" only when the selection looks like
 *  a labelled data range. Keep the heuristic shared across host surfaces. */
export function inferTableHasHeaders(
  workbook: TableHeaderInferenceWorkbook,
  range: Range,
): boolean {
  if (range.r1 <= range.r0) return false;
  if (rangeArea(range) > MAX_TABLE_HEADER_INFERENCE_CELLS) return false;
  let headerTextCount = 0;
  for (let col = range.c0; col <= range.c1; col += 1) {
    const value = workbook.getValue({ sheet: range.sheet, row: range.r0, col });
    if (!isNonEmptyTextValue(value)) return false;
    headerTextCount += 1;
  }
  if (headerTextCount === 0) return false;
  for (let row = range.r0 + 1; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      if (isNonBlankValue(workbook.getValue({ sheet: range.sheet, row, col }))) return true;
    }
  }
  return false;
}

export function customTableStyleId(name: string): string {
  return `${CUSTOM_TABLE_STYLE_PREFIX}${name.trim()}`;
}

export function customPivotTableStyleId(name: string): string {
  return `${CUSTOM_PIVOT_TABLE_STYLE_PREFIX}${name.trim()}`;
}

/** Whether two style names are the same: the desktop app compares them case-insensitively. */
export const sameStyleName = (a: string, b: string): boolean =>
  a.toLocaleLowerCase() === b.toLocaleLowerCase();

/** The gallery id a table or PivotTable style name is shown as: the id of the custom style that
 *  label names (any case), else the name itself, which is a built-in name or '' for none. */
const galleryStyleIdOf = (customs: readonly CustomTableStyle[], name: string): string =>
  customs.find((style) => sameStyleName(style.label, name))?.id ?? name;

/** The table style gallery id the style name `name` is current as. */
export const currentTableStyleId = (
  state: { tables: { customTableStyles?: readonly CustomTableStyle[] } },
  name: string,
): string => galleryStyleIdOf(state.tables.customTableStyles ?? [], name);

/** The PivotTable style gallery id the style name `name` is current as. */
export const currentPivotStyleId = (
  state: { tables: { customPivotTableStyles?: readonly CustomTableStyle[] } },
  name: string,
): string => galleryStyleIdOf(state.tables.customPivotTableStyles ?? [], name);

/** Every built-in table and PivotTable style name. */
const builtinStyleNames = (): string[] => [
  ...BUILTIN_TABLE_STYLES.map((entry) => entry.name),
  ...pivotStyleGallerySections([], { none: false }).flatMap((section) =>
    section.kind === 'custom' ? [] : section.names,
  ),
];

/** Why `name` cannot name a custom style: another custom style of either kind, or a built-in
 *  table or PivotTable style, already has it. `exceptLabel` is the style being modified. */
export function customStyleNameError(
  state: {
    tables: {
      customTableStyles?: readonly CustomTableStyle[];
      customPivotTableStyles?: readonly CustomTableStyle[];
    };
  },
  name: string,
  options: { exceptLabel?: string } = {},
): 'exists' | null {
  const label = name.trim();
  const { exceptLabel } = options;
  const clash = [
    ...(state.tables.customTableStyles ?? []),
    ...(state.tables.customPivotTableStyles ?? []),
  ].some(
    (style) =>
      sameStyleName(style.label, label) &&
      !(exceptLabel !== undefined && sameStyleName(style.label, exceptLabel)),
  );
  return clash || builtinStyleNames().some((builtin) => sameStyleName(builtin, label))
    ? 'exists'
    : null;
}

export function listCustomTableStyles(state: {
  tables: { customTableStyles?: readonly CustomTableStyle[] };
}): readonly CustomTableStyle[] {
  return state.tables.customTableStyles ?? [];
}

export function listCustomPivotTableStyles(state: {
  tables: { customPivotTableStyles?: readonly CustomTableStyle[] };
}): readonly CustomTableStyle[] {
  return state.tables.customPivotTableStyles ?? [];
}

export function customTableStyleById(
  state: { tables: { customTableStyles?: readonly CustomTableStyle[] } },
  id: string,
): CustomTableStyle | null {
  return (state.tables.customTableStyles ?? []).find((style) => style.id === id) ?? null;
}

export function customPivotTableStyleById(
  state: { tables: { customPivotTableStyles?: readonly CustomTableStyle[] } },
  id: string,
): CustomTableStyle | null {
  return (state.tables.customPivotTableStyles ?? []).find((style) => style.id === id) ?? null;
}

export function pivotTableStyleAssignment(
  state: { tables: { pivotTableStyles?: readonly PivotTableStyleAssignment[] } },
  sheetIndex: number,
  pivotIndex: number,
): PivotTableStyleAssignment | null {
  return (
    (state.tables.pivotTableStyles ?? []).find(
      (style) => style.sheetIndex === sheetIndex && style.pivotIndex === pivotIndex,
    ) ?? null
  );
}

function tableVariantFromOptions(input: { banded?: boolean; firstCol?: boolean }): TableVariant {
  if (input.banded && input.firstCol) return 'bandedFirstCol';
  if (input.firstCol) return 'firstCol';
  if (input.banded === false) return 'plain';
  return 'banded';
}

function tableVariantOptions(variant: TableVariant): {
  banded: boolean;
  firstCol: boolean;
} {
  switch (variant) {
    case 'plain':
      return { banded: false, firstCol: false };
    case 'firstCol':
      return { banded: false, firstCol: true };
    case 'bandedFirstCol':
      return { banded: true, firstCol: true };
    default:
      return { banded: true, firstCol: false };
  }
}

export type TableOverlayPatch = Partial<
  Pick<
    TableOverlay,
    | 'range'
    | 'style'
    | 'color'
    | 'showHeader'
    | 'showTotal'
    | 'banded'
    | 'firstCol'
    | 'lastCol'
    | 'colStripes'
    | 'filterButton'
    | 'totals'
    | 'styleName'
    | 'name'
  >
>;

/** The stored overlay `written` became, or null when the store dropped or reverted it. The
 *  store mirror may re-read the totals the engine settled on, and the range it gave a table
 *  whose header or total row was flipped (`engineShaped`), which is not a refusal. */
function storedOverlay(
  store: SpreadsheetStore,
  written: TableOverlay,
  engineShaped = false,
): TableOverlay | null {
  const stored = tableOverlayById(store.getState(), written.id);
  if (stored === written) return stored;
  if (!stored) return null;
  const settled = (t: TableOverlay) => ({ ...t, totals: 0, ...(engineShaped ? { range: 0 } : {}) });
  return JSON.stringify(settled(stored)) === JSON.stringify(settled(written)) ? stored : null;
}

/** Default factory — keeps the construction site small. */
export function defaultTableOverlay(id: string, range: Range): TableOverlay {
  return {
    id,
    source: 'session',
    range,
    style: 'medium',
    color: DEFAULT_TABLE_COLOR,
    showHeader: true,
    showTotal: false,
    banded: true,
  };
}

/** True when `range` would overlap a table other than `id`; tables cannot overlap. */
const overlapsOtherTable = (store: SpreadsheetStore, id: string, range: Range): boolean =>
  store.getState().tables.tables.some((t) => t.id !== id && rangesIntersect(t.range, range));

function defaultTableId(range: Range): string {
  return `table-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}`;
}

/** Apply a Format-as-Table overlay to `range` and return the stored overlay.
 *  With a table-authoring engine the overlay is engine-backed under a fresh
 *  table name; null over another table or when the engine refused it. */
export function formatAsTable(
  store: SpreadsheetStore,
  range: Range,
  options: FormatAsTableOptions = {},
): TableOverlay | null {
  if (blockedByProtection(store, range.sheet, 'formatAsTable')) return null;
  const { workbook, totalsLabel = 'Total', name, ...overlayOptions } = options;
  if (overlapsOtherTable(store, overlayOptions.id ?? defaultTableId(range), range)) return null;
  const engineBacked = workbook?.capabilities?.tableMutate === true;
  const overlay: TableOverlay = {
    ...defaultTableOverlay(overlayOptions.id ?? defaultTableId(range), range),
    ...overlayOptions,
    id: overlayOptions.id ?? defaultTableId(range),
    source: engineBacked ? 'engine' : 'session',
    ...(engineBacked && workbook
      ? { name: name?.trim() || nextTableName(workbook, store.getState().tables.tables) }
      : {}),
    range,
  };
  // A new total row starts with the desktop defaults, the label in the caller's language.
  if (overlay.showTotal && engineBacked && workbook) {
    overlay.totals = defaultTotals(
      workbook,
      range,
      overlay.showHeader,
      range.c1 - range.c0 + 1,
      totalsLabel,
    );
  }
  mutators.upsertTableOverlay(store, overlay);
  return storedOverlay(store, overlay);
}

/** Overlay options for a built-in style id, a custom table style id or a family id; null for an
 *  id that names none of them. */
export function tableStyleOptionsById(
  state: { tables: { customTableStyles?: readonly CustomTableStyle[] } },
  styleId: string,
  color?: string,
  variant: TableVariant = 'banded',
):
  | (Pick<TableOverlay, 'banded' | 'firstCol' | 'styleName'> &
      Partial<Pick<TableOverlay, 'style' | 'color'>>)
  | null {
  const custom = customTableStyleById(state, styleId);
  if (custom) return { styleName: custom.label, ...tableVariantOptions(variant) };
  // A built-in name keeps the name; the grid approximates it by family and hue.
  const builtin = parseTableStyleName(styleId);
  if (builtin) {
    return {
      style: builtin.style,
      color: builtin.color,
      styleName: styleId,
      ...tableVariantOptions(variant),
    };
  }
  if (styleId !== 'light' && styleId !== 'medium' && styleId !== 'dark') return null;
  return { style: styleId, color, styleName: undefined, ...tableVariantOptions(variant) };
}

/** The built-in style a new table takes when the workbook names no default table style. */
const FALLBACK_TABLE_STYLE = 'TableStyleMedium2';

/** The style id a new table takes: the workbook's default table style, else Table Style Medium 2. */
export function defaultTableStyleId(
  state: { tables: { customTableStyles?: readonly CustomTableStyle[] } },
  workbook: Pick<WorkbookHandle, 'getTableStyleDefaults'> | null | undefined,
): string {
  const name = workbook?.getTableStyleDefaults()?.tableStyle ?? '';
  if (name === '') return FALLBACK_TABLE_STYLE;
  if (parseTableStyleName(name)) return name;
  return (
    listCustomTableStyles(state).find((style) => sameStyleName(style.label, name))?.id ??
    FALLBACK_TABLE_STYLE
  );
}

export function formatAsTableByStyleId(
  store: SpreadsheetStore,
  range: Range,
  styleId: string,
  color?: string,
  variant: TableVariant = 'banded',
  options: Pick<
    FormatAsTableOptions,
    'showHeader' | 'showTotal' | 'firstCol' | 'lastCol' | 'workbook' | 'totalsLabel' | 'name'
  > = {},
): TableOverlay | null {
  const look = tableStyleOptionsById(store.getState(), styleId, color, variant);
  return look ? formatAsTable(store, range, { ...look, ...options }) : null;
}

/** Elements and stripe sizes for a new custom style: the caller's, else the active table's look. */
function newStyleBody(
  options: CreateCustomTableStyleOptions,
  source: TableOverlay | null | undefined,
): Pick<CustomTableStyle, 'elements' | 'stripeSizes'> {
  const elements =
    options.elements ??
    galleryStyleElements(
      source?.style ?? 'medium',
      source?.color ?? DEFAULT_TABLE_COLOR,
      tableVariantFromOptions({
        banded: source?.banded ?? true,
        firstCol: source?.firstCol ?? false,
      }),
    );
  return { elements, ...(options.stripeSizes ? { stripeSizes: options.stripeSizes } : {}) };
}

export function createTableStyleFromActiveTable(
  store: SpreadsheetStore,
  history: History | null,
  name: string,
  options: CreateCustomTableStyleOptions = {},
): boolean {
  const label = name.trim();
  if (!label) return false;
  recordTablesChange(history, store, () => {
    const state = store.getState();
    const active = state.selection.active;
    const source = tableOverlayAt(state, active.sheet, active.row, active.col);
    mutators.upsertCustomTableStyle(store, {
      id: customTableStyleId(label),
      label,
      ...newStyleBody(options, source),
    });
  });
  return true;
}

/** Adds a custom PivotTable style; like the table path it assigns the style to no PivotTable. */
export function createPivotTableStyleFromActivePivot(
  store: SpreadsheetStore,
  history: History | null,
  name: string,
  options: CreateCustomTableStyleOptions = {},
): boolean {
  const label = name.trim();
  if (!label) return false;
  recordTablesChange(history, store, () => {
    const state = store.getState();
    const activeTable = tableOverlayAt(
      state,
      state.selection.active.sheet,
      state.selection.active.row,
      state.selection.active.col,
    );
    mutators.upsertCustomPivotTableStyle(store, {
      id: customPivotTableStyleId(label),
      label,
      ...newStyleBody(options, activeTable),
    });
  });
  return true;
}

export function applyPivotTableStyleById(
  store: SpreadsheetStore,
  history: History | null,
  pivot: { sheetIndex: number; pivotIndex: number },
  styleId: string,
): boolean {
  if (!customPivotTableStyleById(store.getState(), styleId)) return false;
  recordTablesChange(history, store, () => {
    mutators.upsertPivotTableStyle(store, { ...pivot, styleId });
  });
  return true;
}

export function listTableOverlays(state: {
  tables: { tables: readonly TableOverlay[] };
}): readonly TableOverlay[] {
  return state.tables.tables;
}

export function sessionTableOverlays(state: {
  tables: { tables: readonly TableOverlay[] };
}): readonly TableOverlay[] {
  return state.tables.tables.filter((t) => t.source === 'session');
}

export function engineTableOverlays(state: {
  tables: { tables: readonly TableOverlay[] };
}): readonly TableOverlay[] {
  return state.tables.tables.filter((t) => t.source === 'engine');
}

export function tableOverlayById(
  state: { tables: { tables: readonly TableOverlay[] } },
  id: string,
): TableOverlay | null {
  return state.tables.tables.find((t) => t.id === id) ?? null;
}

export function tableOverlayAt(
  state: { tables: { tables: readonly TableOverlay[] } },
  sheet: number,
  row: number,
  col: number,
): TableOverlay | null {
  return tableForCell(state.tables.tables, sheet, row, col);
}

/** Why a table cannot take `name`: `invalid` for a name the spreadsheet
 *  forbids, `duplicate` for one another table already uses. */
export type TableNameError = 'invalid' | 'duplicate';

const TABLE_NAME_PATTERN = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u;
const CELL_REFERENCE_PATTERN = /^(?:[A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*|[RrCc])$/;
const MAX_TABLE_NAME_LENGTH = 255;

/** Check `name` as a new name for table `id`; null when it is acceptable. */
export function tableNameError(
  store: SpreadsheetStore,
  id: string,
  name: string,
): TableNameError | null {
  const trimmed = name.trim();
  if (
    !trimmed ||
    trimmed.length > MAX_TABLE_NAME_LENGTH ||
    !TABLE_NAME_PATTERN.test(trimmed) ||
    CELL_REFERENCE_PATTERN.test(trimmed)
  ) {
    return 'invalid';
  }
  const key = trimmed.toLocaleLowerCase();
  const taken = store
    .getState()
    .tables.tables.some((t) => t.id !== id && t.name?.toLocaleLowerCase() === key);
  return taken ? 'duplicate' : null;
}

/** Whether every total cell whose settings `totals` changes on `table` is writable. */
const totalCellsWritable = (
  store: SpreadsheetStore,
  table: TableOverlay,
  totals: readonly TableColumnTotals[] | undefined,
): boolean => {
  if (!table.showTotal || !totals) return false;
  const state = store.getState();
  const { sheet, r1, c0 } = table.range;
  return totals.every(
    (next, i) =>
      JSON.stringify(next) === JSON.stringify(table.totals?.[i]) ||
      isCellWritable(state, { sheet, row: r1, col: c0 + i }),
  );
};

/** Patch a table overlay and return the updated overlay; the store mirror
 *  carries the change to an engine-backed table. Null when the new range
 *  overlaps another table or the engine refused the change. */
export function updateTableOverlay(
  store: SpreadsheetStore,
  id: string,
  patch: TableOverlayPatch,
): TableOverlay | null {
  return patchTableOverlay(store, id, patch);
}

function patchTableOverlay(
  store: SpreadsheetStore,
  id: string,
  patch: TableOverlayPatch,
): TableOverlay | null {
  const current = tableOverlayById(store.getState(), id);
  if (!current) return null;
  // A protected sheet still lets the style toggles change, and a total cell it leaves unlocked.
  const permitted = Object.keys(patch).every(
    (key) =>
      TABLE_FIELDS_ALLOWED_ON_PROTECTED_SHEET.has(key) ||
      (key === 'totals' && totalCellsWritable(store, current, patch.totals)),
  );
  if (!permitted && blockedByProtection(store, current.range.sheet, 'updateTableOverlay'))
    return null;
  if (patch.range && overlapsOtherTable(store, id, patch.range)) return null;
  const { name, ...rest } = patch;
  if (name !== undefined && current.source === 'engine' && tableNameError(store, id, name)) {
    return null;
  }
  const next: TableOverlay = {
    ...current,
    ...rest,
    ...(name !== undefined && current.source === 'engine' ? { name: name.trim() } : {}),
    id: current.id,
    source: current.source,
  };
  mutators.upsertTableOverlay(store, next);
  const flipped = next.showHeader !== current.showHeader || next.showTotal !== current.showTotal;
  return storedOverlay(store, next, flipped);
}

/** Remove a Format-as-Table overlay by id. */
export function clearTable(store: SpreadsheetStore, id: string): void {
  const current = tableOverlayById(store.getState(), id);
  if (current && blockedByProtection(store, current.range.sheet, 'clearTable')) return;
  mutators.removeTableOverlay(store, id);
}

/** Remove every table overlay that intersects `range`. */
export function clearTablesInRange(store: SpreadsheetStore, range: Range): void {
  if (blockedByProtection(store, range.sheet, 'clearTablesInRange')) return;
  mutators.clearTableOverlaysInRange(store, range);
}

/** True when (row, col) sits on the header row of `t`. */
/** True when (row, col) is a header cell that shows its filter button. */
export function showsHeaderFilterButton(t: TableOverlay, row: number, col: number): boolean {
  return t.filterButton !== false && isHeaderRow(t, row, col);
}

export function isHeaderRow(t: TableOverlay, row: number, col: number): boolean {
  if (!t.showHeader) return false;
  if (row !== t.range.r0) return false;
  return col >= t.range.c0 && col <= t.range.c1;
}

/** True when (row, col) is the total row of `t`. */
export function isTotalRow(t: TableOverlay, row: number, col: number): boolean {
  if (!t.showTotal) return false;
  if (row !== t.range.r1) return false;
  return col >= t.range.c0 && col <= t.range.c1;
}

/** True when (row, col) is on the emphasized first data column. */
export function isFirstCol(t: TableOverlay, row: number, col: number): boolean {
  if (!t.firstCol) return false;
  if (col !== t.range.c0) return false;
  if (row < t.range.r0 || row > t.range.r1) return false;
  return true;
}

/** True when (row, col) is on the emphasized last data column. */
export function isLastCol(t: TableOverlay, row: number, col: number): boolean {
  if (!t.lastCol) return false;
  if (col !== t.range.c1) return false;
  if (row < t.range.r0 || row > t.range.r1) return false;
  return true;
}

/** Rows or columns in the first and second stripe band; each defaults to 1. */
export interface StripeBands {
  first?: number;
  second?: number;
}

/** True when offset `n` falls in a first stripe band. */
const inFirstBand = (n: number, { first = 1, second = 1 }: StripeBands): boolean => {
  const a = Math.max(1, first);
  return n % (a + Math.max(1, second)) < a;
};

/** True when the row should paint with the alternate zebra fill. Header
 *  and total rows are excluded. */
export function isBandedRow(
  t: TableOverlay,
  row: number,
  col: number,
  bands: StripeBands = {},
): boolean {
  if (!t.banded) return false;
  if (col < t.range.c0 || col > t.range.c1) return false;
  if (isHeaderRow(t, row, col)) return false;
  if (isTotalRow(t, row, col)) return false;
  if (row < t.range.r0 || row > t.range.r1) return false;
  // The first data row takes the first stripe, as the built-in styles define it.
  const dataStart = t.showHeader ? t.range.r0 + 1 : t.range.r0;
  return inFirstBand(row - dataStart, bands);
}

/** True when the column should paint with the alternate stripe fill. Header
 *  and total rows are excluded. */
export function isBandedCol(
  t: TableOverlay,
  row: number,
  col: number,
  bands: StripeBands = {},
): boolean {
  if (!t.colStripes) return false;
  if (col < t.range.c0 || col > t.range.c1) return false;
  if (row < t.range.r0 || row > t.range.r1) return false;
  if (isHeaderRow(t, row, col) || isTotalRow(t, row, col)) return false;
  return inFirstBand(col - t.range.c0, bands);
}

/** The totals-row cell at (row, col) as its engine table and zero-based column; null elsewhere. */
export function totalsCellAt(
  state: { tables: { tables: readonly TableOverlay[] } },
  sheet: number,
  row: number,
  col: number,
): { table: TableOverlay; column: number } | null {
  const table = tableOverlayAt(state, sheet, row, col);
  if (table?.source !== 'engine' || !isTotalRow(table, row, col)) return null;
  return { table, column: col - table.range.c0 };
}

/** Find the table overlay (if any) that contains a given cell. Tables
 *  are tested in registration order; the first hit wins. */
export function tableForCell(
  tables: readonly TableOverlay[],
  sheet: number,
  row: number,
  col: number,
): TableOverlay | null {
  for (const t of tables) {
    if (t.range.sheet !== sheet) continue;
    if (row < t.range.r0 || row > t.range.r1) continue;
    if (col < t.range.c0 || col > t.range.c1) continue;
    return t;
  }
  return null;
}

/** Add or replace a table overlay (matched by id). Returns a new array. */
export function upsertTable(tables: readonly TableOverlay[], next: TableOverlay): TableOverlay[] {
  const filtered = tables.filter((t) => t.id !== next.id);
  filtered.push(next);
  return filtered;
}

/** Remove a table overlay by id. Returns the same array reference when no
 *  match is found, so callers can short-circuit re-renders. */
export function removeTable(tables: readonly TableOverlay[], id: string): readonly TableOverlay[] {
  const filtered = tables.filter((t) => t.id !== id);
  if (filtered.length === tables.length) return tables;
  return filtered;
}
