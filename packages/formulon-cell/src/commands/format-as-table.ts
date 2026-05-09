import type { Range } from '../engine/types.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

/** UI-only "Format As Table" overlay. Native workbook tables have a full
 * engine model, but this layer can decorate a plain range while writable
 * table APIs are unavailable. */
export type TableStyle = 'light' | 'medium' | 'dark';

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
  /** Render the first row as a header (bold + tinted). Defaults to true. */
  showHeader: boolean;
  /** Render the last row as a total row (bold + tinted). Defaults to false. */
  showTotal: boolean;
  /** Apply zebra fills to data rows. Defaults to true. */
  banded: boolean;
}

export interface FormatAsTableOptions {
  id?: string;
  style?: TableStyle;
  showHeader?: boolean;
  showTotal?: boolean;
  banded?: boolean;
}

export type TableOverlayPatch = Partial<
  Pick<TableOverlay, 'range' | 'style' | 'showHeader' | 'showTotal' | 'banded'>
>;

/** Default factory — keeps the construction site small. */
export function defaultTableOverlay(id: string, range: Range): TableOverlay {
  return {
    id,
    source: 'session',
    range,
    style: 'medium',
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
): TableOverlay {
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
  const next: TableOverlay = { ...current, ...patch, id: current.id, source: 'session' };
  mutators.upsertTableOverlay(store, next);
  return next;
}

/** Remove a session Format-as-Table overlay by id. */
export function clearTable(store: SpreadsheetStore, id: string): void {
  mutators.removeTableOverlay(store, id);
}

/** Remove every session table overlay that intersects `range`. */
export function clearTablesInRange(store: SpreadsheetStore, range: Range): void {
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
