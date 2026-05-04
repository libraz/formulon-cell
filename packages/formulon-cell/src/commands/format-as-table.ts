import type { Range } from '../engine/types.js';

/**
 * UI-only "Format As Table" overlay. Excel's underlying ListObject has a
 * full engine model, but for v0.9 we layer a thin decoration on top of a
 * plain range so users get the visual feedback (header style, zebra rows,
 * total row) without waiting on engine support. xlsx round-trip stays
 * the engine's job; this overlay is session-only.
 */
export type TableStyle = 'light' | 'medium' | 'dark';

export interface TableOverlay {
  /** Stable id used by mutators / pointer routing. */
  id: string;
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

/** Default factory — keeps the construction site small. */
export function defaultTableOverlay(id: string, range: Range): TableOverlay {
  return {
    id,
    range,
    style: 'medium',
    showHeader: true,
    showTotal: false,
    banded: true,
  };
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
