import { MAX_ROW } from '../engine/address.js';
import { totalsFromDetails } from '../engine/table-sync.js';
import { defaultTotals } from '../engine/table-totals-defaults.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { insertCells } from './cell-shift.js';
import {
  type TableColumnTotals,
  type TableNameError,
  type TableOverlay,
  type TableTotalsFunction,
  tableNameError,
  tableOverlayAt,
  tableOverlayById,
  updateTableOverlay,
} from './format-as-table.js';
import type { History } from './history.js';
import { isSheetProtected } from './protection.js';
import { recordTablesChange } from './slice-history.js';

/** Table Design toggles: each flips one show-flag of the active table. */
export type TableOption =
  | 'showHeader'
  | 'showTotal'
  | 'banded'
  | 'firstCol'
  | 'lastCol'
  | 'colStripes'
  | 'filterButton';

/** Table Design ribbon commands and the show-flag each one flips. */
export const TABLE_OPTION_COMMANDS: Readonly<Record<string, TableOption>> = {
  tableHeaderRow: 'showHeader',
  tableTotalRow: 'showTotal',
  tableBandedRows: 'banded',
  tableFirstColumn: 'firstCol',
  tableLastColumn: 'lastCol',
  tableBandedColumns: 'colStripes',
  tableFilterButton: 'filterButton',
};

export type TableRenameResult = 'renamed' | 'unchanged' | 'none' | 'protected' | TableNameError;

/** Totals-row functions the column picker offers, in menu order. */
export const TABLE_TOTALS_FUNCTIONS: readonly TableTotalsFunction[] = [
  'none',
  'average',
  'count',
  'countNums',
  'max',
  'min',
  'sum',
  'stdDev',
  'var',
  'custom',
];

/** The table under the active cell, if any. */
export function activeTableOverlay(state: State): TableOverlay | null {
  const a = state.selection.active;
  return tableOverlayAt(state, a.sheet, a.row, a.col);
}

/** `blocked` means the change would overwrite cells, collide with another table or leave no data row. */
export type TableToggleResult = 'toggled' | 'blocked' | 'none';

const hasTotals = (totals: readonly TableColumnTotals[] | undefined): boolean =>
  !totals || totals.some((t) => t.function !== 'none' || !!t.label);

const bandBlank = (wb: WorkbookHandle, range: Range, row: number): boolean => {
  for (let col = range.c0; col <= range.c1; col += 1) {
    if (wb.getValue({ sheet: range.sheet, row, col }).kind !== 'blank') return false;
  }
  return true;
};

/** Turn the header or total row of an engine table on or off the way the desktop application
 *  does: the ref grows or shrinks by that row, cells below (or the table's own cells, for a
 *  header) shift down when the row is taken, and the cell shift and table change are one undo step. */
const toggleBoundaryRow = (
  store: SpreadsheetStore,
  history: History | null,
  wb: WorkbookHandle,
  table: TableOverlay,
  option: 'showHeader' | 'showTotal',
  totalsLabel: string,
): TableToggleResult => {
  if (option === 'showTotal') return toggleTotalsRow(store, history, wb, table, totalsLabel);
  const r = table.range;
  const turningOn = !table[option];
  const after = { ...table, [option]: turningOn };
  const dataRows =
    r.r1 - r.r0 + 1 + (turningOn ? 1 : -1) - (after.showHeader ? 1 : 0) - (after.showTotal ? 1 : 0);
  if (dataRows < 1) return 'blocked';
  // A header takes an empty row directly above the table; otherwise the table's own cells move
  // down and the header takes its first row. A total row takes the row below, shifting when taken.
  const aboveFree = turningOn && r.r0 > 0 && bandBlank(wb, r, r.r0 - 1);
  const range: Range = turningOn
    ? { ...r, r0: aboveFree ? r.r0 - 1 : r.r0, r1: aboveFree ? r.r1 : r.r1 + 1 }
    : { ...r, r0: r.r0 + 1 };
  const shiftRow = r.r0;
  const needsShift = turningOn && !aboveFree;
  if (turningOn && range.r1 > MAX_ROW) return 'blocked';
  const transaction = history && !history.isReplaying() ? history.begin() : undefined;
  let ok = false;
  try {
    const row: Range = { ...r, r0: shiftRow, r1: shiftRow };
    if (!needsShift || insertCells(store, wb, history, row, 'down')) {
      recordTablesChange(history, store, () => {
        ok = updateTableOverlay(store, table.id, { [option]: turningOn, range }) !== null;
      });
    }
  } finally {
    if (transaction !== undefined) {
      if (ok) history?.end(transaction);
      else history?.abort(transaction);
    }
  }
  return ok ? 'toggled' : 'blocked';
};

/** Turn the total row on or off. The engine grows or shrinks the ref by that row, shifts the
 *  cells below the table's columns (or takes a free row further down) and writes the totals
 *  cells; the table mirror reads the new ref back. */
const toggleTotalsRow = (
  store: SpreadsheetStore,
  history: History | null,
  wb: WorkbookHandle,
  table: TableOverlay,
  totalsLabel: string,
): TableToggleResult => {
  const turningOn = !table.showTotal;
  const r = table.range;
  let ok = false;
  recordTablesChange(history, store, () => {
    const grown = { ...r, r1: r.r1 + 1 };
    const defaults =
      turningOn && !hasTotals(table.totals)
        ? {
            totals: defaultTotals(wb, grown, table.showHeader, r.c1 - r.c0 + 1, totalsLabel),
          }
        : {};
    updateTableOverlay(store, table.id, { showTotal: turningOn, ...defaults });
    // The mirror takes the engine's new ref, or restores the overlay when the engine refuses.
    ok = store.getState().tables.tables.some((t) => t.id === table.id && t.showTotal === turningOn);
  });
  return ok ? 'toggled' : 'blocked';
};

/** Flip a show-flag of the active table. Only engine tables are editable here. */
export function toggleTableOption(
  store: SpreadsheetStore,
  history: History | null,
  wb: WorkbookHandle,
  option: TableOption,
  totalsLabel = 'Total',
): TableToggleResult {
  const table = activeTableOverlay(store.getState());
  if (table?.source !== 'engine') return 'none';
  if (option === 'showHeader' || option === 'showTotal') {
    return toggleBoundaryRow(store, history, wb, table, option, totalsLabel);
  }
  // Filter buttons default to on, so an absent flag reads as shown.
  const on = option === 'filterButton' ? table.filterButton !== false : !!table[option];
  let applied = false;
  recordTablesChange(history, store, () => {
    applied = updateTableOverlay(store, table.id, { [option]: !on }) !== null;
  });
  return applied ? 'toggled' : 'blocked';
}

/** Rename the active engine table; the result says why a name was refused. */
export function renameActiveTable(
  store: SpreadsheetStore,
  history: History | null,
  name: string,
): TableRenameResult {
  const table = activeTableOverlay(store.getState());
  if (table?.source !== 'engine') return 'none';
  const next = name.trim();
  if (next === table.name) return 'unchanged';
  if (isSheetProtected(store.getState(), table.range.sheet)) return 'protected';
  const error = tableNameError(store, table.id, next);
  if (error) return error;
  let renamed = false;
  recordTablesChange(history, store, () => {
    renamed = updateTableOverlay(store, table.id, { name: next }) !== null;
  });
  // The engine refuses a name another object already uses.
  return renamed ? 'renamed' : 'duplicate';
}

/** The engine table's column settings when the overlay has not recorded them yet. */
const engineTotals = (wb: WorkbookHandle, table: TableOverlay): TableColumnTotals[] => {
  const index = table.name ? wb.tableIndexByName(table.name) : null;
  const details = index === null ? null : wb.getTableStyleDetails(index);
  return details ? totalsFromDetails(details) : [];
};

/** Set the totals-row function of one column of table `id`; the store mirror writes it to the
 *  engine, which fills the totals cell. */
export function setTableColumnFunction(
  store: SpreadsheetStore,
  history: History | null,
  workbook: WorkbookHandle,
  id: string,
  column: number,
  fn: TableTotalsFunction,
): boolean {
  const table = tableOverlayById(store.getState(), id);
  if (table?.source !== 'engine') return false;
  const totals = [...(table.totals ?? engineTotals(workbook, table))];
  if (column < 0 || column >= totals.length) return false;
  totals[column] = { function: fn, label: '', formula: '' };
  let applied = false;
  recordTablesChange(history, store, () => {
    applied = updateTableOverlay(store, id, { totals }) !== null;
  });
  return applied;
}
