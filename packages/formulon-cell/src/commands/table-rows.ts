import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { shiftRangeAxis } from './axis-shift.js';
import type { TableOverlay } from './format-as-table.js';
import { shiftFormulaRefs } from './formula-refs.js';
import type { History } from './history.js';
import { recordTablesChange } from './slice-history.js';

/** Table columns a partial cell shift is confined to; null for whole sheet rows. */
export interface TableBand {
  c0: number;
  c1: number;
}

/**
 * Re-point the tables of `sheet` after rows were inserted (`delta > 0`) or
 * deleted (`delta < 0`) at row `split`. A table below the edit moves, one the
 * edit lands inside (or on its totals row) grows or shrinks, and one whose
 * header row is deleted goes away. A new data row takes the calculated
 * columns of the rows beside it.
 *
 * With a `band` (Insert/Delete Cells shifting up or down), only a band that
 * covers exactly the table's columns inside its data rows edits the table.
 */
export function shiftTablesForRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  split: number,
  delta: number,
  band: TableBand | null = null,
): void {
  const tables = store.getState().tables.tables.filter((t) => t.range.sheet === sheet);
  if (tables.length === 0 || delta === 0) return;
  const count = Math.abs(delta);
  const edits: { table: TableOverlay; range: ReturnType<typeof shiftRangeAxis> }[] = [];
  for (const table of tables) {
    const t = table.range;
    if (band) {
      const dataEnd = table.showTotal ? t.r1 - 1 : t.r1;
      const inBody =
        band.c0 === t.c0 &&
        band.c1 === t.c1 &&
        split > t.r0 &&
        (delta > 0 ? split <= t.r1 : split + count - 1 <= dataEnd);
      if (!inBody) continue;
    }
    if (delta < 0 && table.showHeader && split <= t.r0 && split + count - 1 >= t.r0) {
      edits.push({ table, range: null });
      continue;
    }
    const range = shiftRangeAxis(t, 'row', split, delta);
    if (range) {
      if (range.r0 !== t.r0 || range.r1 !== t.r1) edits.push({ table, range });
    } else edits.push({ table, range: null });
  }
  if (edits.length === 0) return;
  recordTablesChange(history, store, () => {
    for (const { table, range } of edits) {
      if (range) mutators.upsertTableOverlay(store, { ...table, range });
      else mutators.removeTableOverlay(store, table.id);
    }
  });
  if (delta < 0) return;
  wb.withBatchedRecalc(() => {
    for (const { table, range } of edits) {
      const t = table.range;
      // The edit landed inside the table, so its new rows are data rows.
      if (range && split > t.r0 && split <= t.r1) {
        fillCalculatedColumns(wb, range, table, [split, split + count - 1]);
      }
    }
  });
}

/** Give each new data row the formula a column holds in every row beside it. */
function fillCalculatedColumns(
  wb: WorkbookHandle,
  range: NonNullable<TableOverlay['range']>,
  table: TableOverlay,
  [first, last]: [number, number],
): void {
  const start = range.r0 + (table.showHeader ? 1 : 0);
  const end = table.showTotal ? range.r1 - 1 : range.r1;
  const existing: number[] = [];
  for (let row = start; row <= end; row += 1) if (row < first || row > last) existing.push(row);
  const templateRow = existing.includes(first - 1) ? first - 1 : existing[0];
  if (templateRow === undefined) return;
  for (let col = range.c0; col <= range.c1; col += 1) {
    const formula = wb.cellFormula({ sheet: range.sheet, row: templateRow, col });
    if (!formula) continue;
    const isCalculated = existing.every(
      (row) =>
        wb.cellFormula({ sheet: range.sheet, row, col }) ===
        shiftFormulaRefs(formula, row - templateRow, 0),
    );
    if (!isCalculated) continue;
    for (let row = first; row <= last; row += 1) {
      wb.setFormula(
        { sheet: range.sheet, row, col },
        shiftFormulaRefs(formula, row - templateRow, 0),
      );
    }
  }
}
