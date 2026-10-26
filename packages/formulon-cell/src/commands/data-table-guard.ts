import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type DataTableInfo, formulaCallsTable } from '../engine/workbook-handle-data-tables.js';
import { selectionCoversRange } from '../store/selection-geometry.js';
import type { SelectionSlice } from '../store/types.js';

/** Stop-level alert shown when an edit would touch part of a data table. */
export interface DataTableGuardAlert {
  severity: 'stop';
  message: string;
}

type DataTableGuardSink = (alert: DataTableGuardAlert) => void;

/** Alert wording, in the UI language. */
export interface DataTableGuardMessages {
  /** An edit reaches a data table's results. */
  cannotChange: string;
  /** A typed formula calls TABLE(), which only Data Table may create. */
  tableFunction: string;
}

const ENGLISH_MESSAGES: DataTableGuardMessages = {
  cannotChange: 'Cannot change part of a data table.',
  tableFunction: "That function isn't valid.",
};

const guards = new WeakMap<
  WorkbookHandle,
  { sink: DataTableGuardSink; messages: DataTableGuardMessages }
>();

/** Route `blockDataTableEdit` alerts for `wb` to the host's alert surface. */
export function setDataTableGuardSink(
  wb: WorkbookHandle,
  sink: DataTableGuardSink | null,
  messages: DataTableGuardMessages = ENGLISH_MESSAGES,
): void {
  if (sink) guards.set(wb, { sink, messages });
  else guards.delete(wb);
}

const messagesFor = (wb: WorkbookHandle): DataTableGuardMessages =>
  guards.get(wb)?.messages ?? ENGLISH_MESSAGES;

export function dataTableGuardAlert(wb: WorkbookHandle): DataTableGuardAlert {
  return { severity: 'stop', message: messagesFor(wb).cannotChange };
}

/** The alert for a typed TABLE() formula. */
export function tableFunctionAlert(wb: WorkbookHandle): DataTableGuardAlert {
  return { severity: 'stop', message: messagesFor(wb).tableFunction };
}

/** Whether a data table's result block holds any cell of `range`. */
export function dataTableOccupies(wb: WorkbookHandle, range: Range): boolean {
  if (!wb.capabilities?.dataTables) return false;
  if (range.r0 === range.r1 && range.c0 === range.c1) {
    return wb.dataTableAt({ sheet: range.sheet, row: range.r0, col: range.c0 }) !== null;
  }
  return wb.dataTablesInRange(range).length > 0;
}

/** Show `alert` on the host's alert surface. */
export function alertDataTableGuard(wb: WorkbookHandle, alert: DataTableGuardAlert): void {
  guards.get(wb)?.sink(alert);
}

/** Tell the user an edit would change part of a data table. */
export function alertDataTableEdit(wb: WorkbookHandle): void {
  alertDataTableGuard(wb, dataTableGuardAlert(wb));
}

/** Reject a write over any of `ranges` that touches a data table's results:
 *  alerts the user and returns true when the caller must leave the sheet
 *  untouched. */
export function blockDataTableEdit(wb: WorkbookHandle, ...ranges: Range[]): boolean {
  if (!ranges.some((range) => dataTableOccupies(wb, range))) return false;
  alertDataTableEdit(wb);
  return true;
}

/** How a selection meets the data tables of its sheet. */
export interface DataTableSelectionPlan {
  /** Tables whose whole result block is selected; Delete removes these. */
  whole: DataTableInfo[];
  /** A result block is only partly selected, so the edit must be refused. */
  partial: boolean;
}

export function planDataTableSelection(
  wb: WorkbookHandle,
  selection: Pick<SelectionSlice, 'range' | 'extraRanges'>,
): DataTableSelectionPlan {
  const plan: DataTableSelectionPlan = { whole: [], partial: false };
  if (!wb.capabilities?.dataTables) return plan;
  const seen = new Set<string>();
  for (const range of [selection.range, ...(selection.extraRanges ?? [])]) {
    for (const table of wb.dataTablesInRange(range)) {
      const key = `${table.range.sheet}:${table.range.r0}:${table.range.c0}`;
      if (seen.has(key)) continue;
      seen.add(key);
      if (selectionCoversRange(selection, table.range)) plan.whole.push(table);
      else plan.partial = true;
    }
  }
  return plan;
}

/** Refuse a write to any of `cells` that lies in a data table's results:
 *  alerts and returns true when the caller must write nothing. */
export function blockDataTableCells(wb: WorkbookHandle, cells: readonly Addr[]): boolean {
  if (!dataTableCellAmong(wb, cells)) return false;
  alertDataTableEdit(wb);
  return true;
}

/** Whether inserting or deleting whole rows/columns would split a data table.
 *  A table spans its result block plus the formula row above and the input
 *  column left of it; an insert must fall outside that span or on its first
 *  line, and a delete must miss it or take it whole. */
export function axisEditSplitsDataTable(
  wb: WorkbookHandle,
  sheet: number,
  axis: 'row' | 'col',
  kind: 'insert' | 'delete',
  at: number,
  count: number,
): boolean {
  if (!wb.capabilities?.dataTables) return false;
  // Only a table whose results reach the edited lines (or the line below a
  // deleted band, for its formula row or input column) can be split.
  const lastLine = Math.min(
    axis === 'row' ? MAX_ROW : MAX_COL,
    at + (kind === 'delete' ? count : 0),
  );
  const lines: Range =
    axis === 'row'
      ? { sheet, r0: at, r1: lastLine, c0: 0, c1: MAX_COL }
      : { sheet, r0: 0, r1: MAX_ROW, c0: at, c1: lastLine };
  for (const table of wb.dataTablesInRange(lines)) {
    const first = (axis === 'row' ? table.range.r0 : table.range.c0) - 1;
    const last = axis === 'row' ? table.range.r1 : table.range.c1;
    if (kind === 'insert') {
      if (at > first && at <= last) return true;
      continue;
    }
    const end = at + count - 1;
    const overlaps = at <= last && end >= first;
    const covers = at <= first && end >= last;
    if (overlaps && !covers) return true;
  }
  return false;
}

/** The first of `cells` lying in a data table's result block, or `null`. One
 *  range scan per sheet, so a large batch costs no per-cell engine lookups. */
export function dataTableCellAmong(wb: WorkbookHandle, cells: readonly Addr[]): Addr | null {
  if (!wb.capabilities?.dataTables || cells.length === 0) return null;
  const bounds = new Map<number, Range>();
  for (const { sheet, row, col } of cells) {
    const b = bounds.get(sheet);
    bounds.set(
      sheet,
      b
        ? {
            sheet,
            r0: Math.min(b.r0, row),
            c0: Math.min(b.c0, col),
            r1: Math.max(b.r1, row),
            c1: Math.max(b.c1, col),
          }
        : { sheet, r0: row, c0: col, r1: row, c1: col },
    );
  }
  const tables = [...bounds.values()].flatMap((range) => wb.dataTablesInRange(range));
  if (tables.length === 0) return null;
  return (
    cells.find((cell) =>
      tables.some(
        ({ range }) =>
          range.sheet === cell.sheet &&
          cell.row >= range.r0 &&
          cell.row <= range.r1 &&
          cell.col >= range.c0 &&
          cell.col <= range.c1,
      ),
    ) ?? null
  );
}

/** Refuse a cell shift over `band` that would move any part of a data table —
 *  its results, formula row or input column. Alerts and returns true. */
export function blockDataTableShift(wb: WorkbookHandle, band: Range): boolean {
  if (!wb.capabilities?.dataTables) return false;
  // A table's span reaches one row above and one column left of its results,
  // so results just past the band's bottom or right edge still count.
  const reach: Range = {
    ...band,
    r1: Math.min(MAX_ROW, band.r1 + 1),
    c1: Math.min(MAX_COL, band.c1 + 1),
  };
  const moved = wb
    .dataTablesInRange(reach)
    .some(
      ({ range }) =>
        range.r0 - 1 <= band.r1 &&
        range.r1 >= band.r0 &&
        range.c0 - 1 <= band.c1 &&
        range.c1 >= band.c0,
    );
  if (moved) alertDataTableEdit(wb);
  return moved;
}

/** Refuse a paste or import whose formulas call TABLE(): alerts and returns
 *  true when the caller must write nothing. */
export function blockTableFunction(
  wb: WorkbookHandle,
  formulas: Iterable<string | null | undefined>,
): boolean {
  for (const formula of formulas) {
    if (!formulaCallsTable(formula)) continue;
    alertDataTableGuard(wb, tableFunctionAlert(wb));
    return true;
  }
  return false;
}
