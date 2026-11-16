import { tableRange } from './range-resolver.js';
import type { Addr, CellValue, Range } from './types.js';

/** A cell's content as the occupancy test reads it. */
export type SpillCellAt = (
  row: number,
  col: number,
) => { readonly value: CellValue; readonly formula: string | null } | undefined;

/** The workbook reads that name the objects claiming cells on a sheet. */
export interface SpillOccupantSource {
  getMerges(sheet: number): Range[];
  getTables(): readonly { readonly sheetIndex: number; readonly ref: string }[];
  pivotBounds(sheet: number): { pivotIndex: number; range: Range }[];
  dataTables(sheet: number): readonly { readonly range: Range }[];
}

const inRange = (r: Range, row: number, col: number): boolean =>
  row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

/**
 * Whether a cell of `sheet` holds a spill or a PivotTable report back: a value or formula
 * (`=""` included; spill, legacy array and data-table results carry values), or any cell of a
 * merge, a table, a data table or another report. A format alone does not.
 * `ownPivot` leaves that report's own extent out.
 */
export function spillOccupancy(
  src: SpillOccupantSource,
  sheet: number,
  cellAt: SpillCellAt,
  ownPivot?: number,
): (row: number, col: number) => boolean {
  let regions: Range[] | null = null;
  const claimed = (): Range[] => {
    regions ??= [
      ...src.getMerges(sheet),
      ...src
        .getTables()
        .filter((t) => t.sheetIndex === sheet)
        .map((t) => tableRange(sheet, t.ref))
        .filter((r): r is Range => r !== null),
      ...src
        .pivotBounds(sheet)
        .filter((p) => p.pivotIndex !== ownPivot)
        .map((p) => p.range),
      ...src.dataTables(sheet).map((t) => t.range),
    ];
    return regions;
  };
  return (row, col) => {
    const cell = cellAt(row, col);
    if (cell && (cell.formula !== null || cell.value.kind !== 'blank')) return true;
    return claimed().some((r) => inRange(r, row, col));
  };
}

/** Rows and columns to read, in order; a renderer passes its visible axes. */
interface SpillAxes {
  readonly rows: readonly number[];
  readonly cols: readonly number[];
}

const span = (from: number, to: number): number[] =>
  Array.from({ length: Math.max(0, to - from + 1) }, (_, i) => from + i);

/** The cells of `target` that hold back the spill anchored at `anchor`, row-major. With `axes`,
 *  only the target cells on those rows and columns are read. */
export function spillBlockers(
  target: Range,
  anchor: Addr,
  occupied: (row: number, col: number) => boolean,
  axes: SpillAxes = { rows: span(target.r0, target.r1), cols: span(target.c0, target.c1) },
): Addr[] {
  const out: Addr[] = [];
  const cols = axes.cols.filter((col) => col >= target.c0 && col <= target.c1);
  for (const row of axes.rows) {
    if (row < target.r0 || row > target.r1) continue;
    for (const col of cols) {
      if (row === anchor.row && col === anchor.col) continue;
      if (occupied(row, col)) out.push({ sheet: target.sheet, row, col });
    }
  }
  return out;
}
