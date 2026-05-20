import type { CellValue, Range } from '../engine/types.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { type History, recordTablesChange } from './history.js';
import { isSheetProtected } from './protection.js';

/** UI-only "Format As Table" overlay. Native workbook tables have a full
 * engine model, but this layer can decorate a plain range while writable
 * table APIs are unavailable. */
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
export const DEFAULT_TABLE_COLOR = '#5b9bd5';

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

export interface TableOverlay {
  /** Stable id used by mutators / pointer routing. */
  id: string;
  /** Source of the overlay. Loaded workbook tables are engine-backed/read-only;
   *  session tables are visual authoring overlays created by the UI. */
  source: 'engine' | 'session';
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
}

export interface TableHeaderInferenceWorkbook {
  getValue(addr: { sheet: number; row: number; col: number }): CellValue;
}

const CUSTOM_TABLE_STYLE_PREFIX = 'custom-table:';
const CUSTOM_PIVOT_TABLE_STYLE_PREFIX = 'custom-pivot-table:';

const isNonEmptyTextValue = (value: CellValue): boolean =>
  value.kind === 'text' && value.value.trim().length > 0;

const isNonBlankValue = (value: CellValue): boolean => value.kind !== 'blank';

/** Excel pre-checks "My table has headers" only when the selection looks like
 *  a labelled data range. Keep the heuristic shared across host surfaces. */
export function inferTableHasHeaders(
  workbook: TableHeaderInferenceWorkbook,
  range: Range,
): boolean {
  if (range.r1 <= range.r0) return false;
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
    'range' | 'style' | 'color' | 'showHeader' | 'showTotal' | 'banded' | 'firstCol' | 'lastCol'
  >
>;

const blockedByProtection = (store: SpreadsheetStore, sheet: number, op: string): boolean => {
  if (!isSheetProtected(store.getState(), sheet)) return false;
  // eslint-disable-next-line no-console
  console.warn(`formulon-cell: ${op} blocked — sheet ${sheet} is protected`);
  return true;
};

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

function defaultTableId(range: Range): string {
  return `table-${range.sheet}-${range.r0}-${range.c0}-${range.r1}-${range.c1}`;
}

/** Apply a session Format-as-Table overlay to `range` and return the stored
 *  overlay. This stays UI-level until the engine exposes writable table APIs. */
export function formatAsTable(
  store: SpreadsheetStore,
  range: Range,
  options: FormatAsTableOptions = {},
): TableOverlay | null {
  if (blockedByProtection(store, range.sheet, 'formatAsTable')) return null;
  const overlay: TableOverlay = {
    ...defaultTableOverlay(options.id ?? defaultTableId(range), range),
    ...options,
    id: options.id ?? defaultTableId(range),
    source: 'session',
    range,
  };
  mutators.upsertTableOverlay(store, overlay);
  return overlay;
}

export function formatAsTableByStyleId(
  store: SpreadsheetStore,
  range: Range,
  styleId: string,
  color?: string,
  variant: CustomTableStyle['variant'] = 'banded',
  options: Pick<FormatAsTableOptions, 'showHeader' | 'showTotal' | 'firstCol' | 'lastCol'> = {},
): TableOverlay | null {
  const custom = customTableStyleById(store.getState(), styleId);
  if (custom) {
    return formatAsTable(store, range, {
      style: custom.style,
      color: custom.color,
      ...tableVariantOptions(custom.variant),
      ...options,
    });
  }
  return formatAsTable(store, range, {
    style: styleId as TableStyle,
    color,
    ...tableVariantOptions(variant),
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

/** Patch a session table overlay and return the updated overlay. Engine-backed
 *  overlays are intentionally read-only at this layer. */
export function updateTableOverlay(
  store: SpreadsheetStore,
  id: string,
  patch: TableOverlayPatch,
): TableOverlay | null {
  const current = tableOverlayById(store.getState(), id);
  if (!current || current.source !== 'session') return null;
  if (blockedByProtection(store, current.range.sheet, 'updateTableOverlay')) return null;
  const next: TableOverlay = { ...current, ...patch, id: current.id, source: 'session' };
  mutators.upsertTableOverlay(store, next);
  return next;
}

/** Remove a session Format-as-Table overlay by id. */
export function clearTable(store: SpreadsheetStore, id: string): void {
  const current = tableOverlayById(store.getState(), id);
  if (current && blockedByProtection(store, current.range.sheet, 'clearTable')) return;
  mutators.removeTableOverlay(store, id);
}

/** Remove every session table overlay that intersects `range`. */
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
  // First data row is "even" — paint zebra on every other row from there.
  const dataStart = t.showHeader ? t.range.r0 + 1 : t.range.r0;
  return ((row - dataStart) & 1) === 1;
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
