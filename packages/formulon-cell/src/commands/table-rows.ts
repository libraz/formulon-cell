import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { shiftRangeAxis } from './axis-shift.js';
import type { TableOverlay } from './format-as-table.js';
import { shiftFormulaRefs } from './formula-refs.js';
import type { History } from './history.js';
import { recordTablesChange, recordTablesFollowingEngine } from './slice-history.js';

/**
 * Re-point the tables of `sheet` after rows were inserted (`delta > 0`) or
 * deleted (`delta < 0`) at row `split`. A table below the edit moves and one
 * the edit lands inside grows or shrinks; deleting its totals row turns the
 * totals off. A new data row takes the calculated columns of the rows beside
 * it. A delete never reaches a header row (`rowDeleteHitsTableHeader`). When the engine ran the
 * edit (`engineMoved`) it moved its tables already, so the overlays only follow it.
 */
export function shiftTablesForRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  split: number,
  delta: number,
  engineMoved: boolean,
): void {
  const tables = store.getState().tables.tables.filter((t) => t.range.sheet === sheet);
  if (tables.length === 0 || delta === 0) return;
  const count = Math.abs(delta);
  const edits: { table: TableOverlay; range: ReturnType<typeof shiftRangeAxis> }[] = [];
  for (const table of tables) {
    const t = table.range;
    const range = shiftRangeAxis(t, 'row', split, delta);
    if (range) {
      const totalsGone = delta < 0 && table.showTotal && split <= t.r1 && split + count - 1 >= t.r1;
      if (range.r0 !== t.r0 || range.r1 !== t.r1 || totalsGone) {
        edits.push({ table: totalsGone ? { ...table, showTotal: false } : table, range });
      }
    } else edits.push({ table, range: null });
  }
  if (edits.length === 0) return;
  const shift = (): void => {
    for (const { table, range } of edits) {
      if (range) mutators.upsertTableOverlay(store, { ...table, range });
      else mutators.removeTableOverlay(store, table.id);
    }
  };
  if (engineMoved) recordTablesFollowingEngine(history, store, wb, shift);
  else recordTablesChange(history, store, shift);
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

/** Whether deleting `count` rows from `at` takes in the header row of a table on `sheet`. */
export function rowDeleteHitsTableHeader(
  store: SpreadsheetStore,
  sheet: number,
  at: number,
  count: number,
): boolean {
  return store
    .getState()
    .tables.tables.some(
      (t) =>
        t.range.sheet === sheet && t.showHeader && at <= t.range.r0 && at + count - 1 >= t.range.r0,
    );
}

/**
 * The first data row of a table whose data rows the deletion of `count` rows
 * from `at` would all remove. The engine keeps that row, blank, and removes
 * only the rows below it.
 */
export function keptTableDataRow(
  store: SpreadsheetStore,
  sheet: number,
  at: number,
  count: number,
): { range: NonNullable<TableOverlay['range']>; row: number } | null {
  for (const table of store.getState().tables.tables) {
    const t = table.range;
    if (t.sheet !== sheet) continue;
    const first = t.r0 + (table.showHeader ? 1 : 0);
    const last = table.showTotal ? t.r1 - 1 : t.r1;
    if (at === first && at + count - 1 >= last) return { range: t, row: first };
  }
  return null;
}

/**
 * Re-point the tables of `sheet` after `count` columns were inserted at
 * `split`: the insert widens the table it lands inside and moves one to its
 * right. When the engine ran the insert (`engineMoved`) the overlays only follow it.
 */
export function shiftTablesForCols(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  split: number,
  count: number,
  engineMoved: boolean,
): void {
  const edits: { table: TableOverlay; range: NonNullable<ReturnType<typeof shiftRangeAxis>> }[] =
    [];
  for (const table of store.getState().tables.tables) {
    const t = table.range;
    if (t.sheet !== sheet) continue;
    const range = shiftRangeAxis(t, 'col', split, count);
    if (range && (range.c0 !== t.c0 || range.c1 !== t.c1)) edits.push({ table, range });
  }
  if (edits.length === 0) return;
  const shift = (): void => {
    for (const { table, range } of edits) mutators.upsertTableOverlay(store, { ...table, range });
  };
  if (engineMoved) recordTablesFollowingEngine(history, store, wb, shift);
  else recordTablesChange(history, store, shift);
}
