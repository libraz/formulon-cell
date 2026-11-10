import type { Addr, CellValue } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { shiftRangeAxis } from './axis-shift.js';
import type { TableOverlay } from './format-as-table.js';
import { shiftFormulaRefs } from './formula-refs.js';
import type { History } from './history.js';
import { recordTablesChange } from './slice-history.js';

/**
 * Re-point the tables of `sheet` after rows were inserted (`delta > 0`) or
 * deleted (`delta < 0`) at row `split`. A table below the edit moves and one
 * the edit lands inside grows or shrinks; deleting its totals row turns the
 * totals off. A new data row takes the calculated columns of the rows beside
 * it. A delete never reaches a header row (`rowDeleteHitsTableHeader`).
 */
export function shiftTablesForRows(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  history: History | null,
  sheet: number,
  split: number,
  delta: number,
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
 * from `at` would all remove. Excel keeps that row, blank, instead.
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

/** Blank `row` across the columns of `range`, recording the cells for undo. */
export function clearTableRow(
  wb: WorkbookHandle,
  history: History | null,
  range: NonNullable<TableOverlay['range']>,
  row: number,
): void {
  const before: { addr: Addr; value: CellValue; formula: string | null }[] = [];
  for (let col = range.c0; col <= range.c1; col += 1) {
    const addr = { sheet: range.sheet, row, col };
    before.push({ addr, value: wb.getValue(addr), formula: wb.cellFormula(addr) });
  }
  const clear = (): void => {
    wb.withBatchedRecalc(() => {
      for (const { addr } of before) wb.setBlank(addr);
    });
    wb.recalcAuto();
  };
  clear();
  if (!history || history.isReplaying()) return;
  history.push({
    undo: () => {
      wb.withBatchedRecalc(() => {
        for (const { addr, value, formula } of before) writeCell(wb, addr, value, formula);
      });
      wb.recalcAuto();
    },
    redo: clear,
  });
}

/**
 * Re-point the tables of `sheet` after `count` columns were inserted at
 * `split`: the insert widens the table it lands inside and moves one to its
 * right.
 */
export function shiftTablesForCols(
  store: SpreadsheetStore,
  history: History | null,
  sheet: number,
  split: number,
  count: number,
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
  recordTablesChange(history, store, () => {
    for (const { table, range } of edits) mutators.upsertTableOverlay(store, { ...table, range });
  });
}
