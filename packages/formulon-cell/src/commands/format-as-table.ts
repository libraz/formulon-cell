import { defaultTotals } from '../engine/table-totals-defaults.js';
import type { CellValue, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeArea, rangesIntersect } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { blockedByProtection } from './protection.js';
import { recordTablesChange } from './slice-history.js';

/** "Format As Table" overlay. With a table-authoring engine every overlay is
 * backed by an engine table of the same `name`, which the store mirror keeps
 * in step; without one it decorates a plain range for the session only. */
export type TableStyle = 'light' | 'medium' | 'dark';

/** Built-in table-style hues — one column per swatch in the style gallery,
 *  ordered to mirror the desktop spreadsheet's themed accent columns. */
export const TABLE_STYLE_COLORS: readonly string[] = [
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

export interface CustomTableStyle {
  id: string;
  label: string;
  style: TableStyle;
  color?: string;
  variant: 'plain' | 'banded' | 'firstCol' | 'bandedFirstCol';
}

export interface PivotTableStyleAssignment {
  sheetIndex: number;
  pivotIndex: number;
  styleId: string;
}

export interface CreateCustomTableStyleOptions {
  style?: TableStyle;
  color?: string;
  variant?: CustomTableStyle['variant'];
}

/** Derive the header / banded-row fills for a `(style, color)` pair. */
export function tableStyleSwatch(
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

/** Unique column names for a table over `range`, from its header row when present. */
export const tableColumns = (
  workbook: Pick<WorkbookHandle, 'getValue'>,
  range: Range,
  hasHeaders: boolean,
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
        : `Column${col - range.c0 + 1}`;
    let name = base;
    let suffix = 2;
    while (used.has(name.toLocaleLowerCase())) name = `${base}${suffix++}`;
    used.add(name.toLocaleLowerCase());
    columns.push(name);
  }
  return columns;
};

const nextTableName = (workbook: WorkbookHandle, overlays: readonly TableOverlay[]): string => {
  const existing = new Set(workbook.getTables().map((table) => table.name.toLocaleLowerCase()));
  for (const overlay of overlays) if (overlay.name) existing.add(overlay.name.toLocaleLowerCase());
  let n = 1;
  while (existing.has(`table${n}`)) n += 1;
  return `Table${n}`;
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

export function tableVariantFromOptions(input: {
  banded?: boolean;
  firstCol?: boolean;
}): CustomTableStyle['variant'] {
  if (input.banded && input.firstCol) return 'bandedFirstCol';
  if (input.firstCol) return 'firstCol';
  if (input.banded === false) return 'plain';
  return 'banded';
}

export function tableVariantOptions(variant: CustomTableStyle['variant']): {
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
    | 'totals'
    | 'styleName'
    | 'name'
  >
>;

/** The stored overlay `written` became, or null when the store dropped or reverted it. The
 *  store mirror may re-read the totals the engine settled on, which is not a refusal. */
function storedOverlay(store: SpreadsheetStore, written: TableOverlay): TableOverlay | null {
  const stored = tableOverlayById(store.getState(), written.id);
  if (stored === written) return stored;
  if (!stored) return null;
  return JSON.stringify({ ...stored, totals: 0 }) === JSON.stringify({ ...written, totals: 0 })
    ? stored
    : null;
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
  const { workbook, totalsLabel = 'Total', ...overlayOptions } = options;
  if (overlapsOtherTable(store, overlayOptions.id ?? defaultTableId(range), range)) return null;
  const engineBacked = workbook?.capabilities?.tableMutate === true;
  const overlay: TableOverlay = {
    ...defaultTableOverlay(overlayOptions.id ?? defaultTableId(range), range),
    ...overlayOptions,
    id: overlayOptions.id ?? defaultTableId(range),
    source: engineBacked ? 'engine' : 'session',
    ...(engineBacked && workbook
      ? { name: nextTableName(workbook, store.getState().tables.tables) }
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

/** Overlay options for a built-in style id or a custom table style id. */
export function tableStyleOptionsById(
  state: { tables: { customTableStyles?: readonly CustomTableStyle[] } },
  styleId: string,
  color?: string,
  variant: CustomTableStyle['variant'] = 'banded',
): Pick<TableOverlay, 'style' | 'color' | 'banded' | 'firstCol' | 'styleName'> {
  const custom = customTableStyleById(state, styleId);
  if (custom) {
    return {
      style: custom.style,
      color: custom.color,
      styleName: custom.label,
      ...tableVariantOptions(custom.variant),
    };
  }
  return {
    style: styleId as TableStyle,
    color,
    styleName: undefined,
    ...tableVariantOptions(variant),
  };
}

export function formatAsTableByStyleId(
  store: SpreadsheetStore,
  range: Range,
  styleId: string,
  color?: string,
  variant: CustomTableStyle['variant'] = 'banded',
  options: Pick<
    FormatAsTableOptions,
    'showHeader' | 'showTotal' | 'firstCol' | 'lastCol' | 'workbook' | 'totalsLabel'
  > = {},
): TableOverlay | null {
  return formatAsTable(store, range, {
    ...tableStyleOptionsById(store.getState(), styleId, color, variant),
    ...options,
  });
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
      style: options.style ?? source?.style ?? 'medium',
      color: options.color ?? source?.color ?? DEFAULT_TABLE_COLOR,
      variant:
        options.variant ??
        tableVariantFromOptions({
          banded: source?.banded ?? true,
          firstCol: source?.firstCol ?? false,
        }),
    });
  });
  return true;
}

export function createPivotTableStyleFromActivePivot(
  store: SpreadsheetStore,
  history: History | null,
  name: string,
  pivot?: { sheetIndex: number; pivotIndex: number } | null,
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
    const style = {
      id: customPivotTableStyleId(label),
      label,
      style: options.style ?? activeTable?.style ?? 'medium',
      color: options.color ?? activeTable?.color ?? DEFAULT_TABLE_COLOR,
      variant:
        options.variant ??
        tableVariantFromOptions({
          banded: activeTable?.banded ?? true,
          firstCol: activeTable?.firstCol ?? false,
        }),
    };
    mutators.upsertCustomPivotTableStyle(store, style);
    if (pivot) {
      mutators.upsertPivotTableStyle(store, {
        sheetIndex: pivot.sheetIndex,
        pivotIndex: pivot.pivotIndex,
        styleId: style.id,
      });
    }
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

/** Patch a table overlay and return the updated overlay; the store mirror
 *  carries the change to an engine-backed table. Null when the new range
 *  overlaps another table or the engine refused the change. */
export function updateTableOverlay(
  store: SpreadsheetStore,
  id: string,
  patch: TableOverlayPatch,
): TableOverlay | null {
  const current = tableOverlayById(store.getState(), id);
  if (!current) return null;
  if (blockedByProtection(store, current.range.sheet, 'updateTableOverlay')) return null;
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
  return storedOverlay(store, next);
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

/** True when the row should paint with the alternate zebra fill. Header
 *  and total rows are excluded. */
export function isBandedRow(t: TableOverlay, row: number, col: number): boolean {
  if (!t.banded) return false;
  if (col < t.range.c0 || col > t.range.c1) return false;
  if (isHeaderRow(t, row, col)) return false;
  if (isTotalRow(t, row, col)) return false;
  if (row < t.range.r0 || row > t.range.r1) return false;
  // The first data row takes the first stripe, as the built-in styles define it.
  const dataStart = t.showHeader ? t.range.r0 + 1 : t.range.r0;
  return ((row - dataStart) & 1) === 0;
}

/** True when the column should paint with the alternate stripe fill. Header
 *  and total rows are excluded. */
export function isBandedCol(t: TableOverlay, row: number, col: number): boolean {
  if (!t.colStripes) return false;
  if (col < t.range.c0 || col > t.range.c1) return false;
  if (row < t.range.r0 || row > t.range.r1) return false;
  if (isHeaderRow(t, row, col) || isTotalRow(t, row, col)) return false;
  return ((col - t.range.c0) & 1) === 0;
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
