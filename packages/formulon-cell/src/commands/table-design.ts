import { totalsFromDetails } from '../engine/table-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore, State } from '../store/store.js';
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

/** Turn the header or total row of an engine table on or off. Only the flag is patched: the engine
 *  grows or shrinks the ref, moves or writes the cells, and restores a total row's saved column
 *  settings; the table mirror reads the new ref back, or restores the overlay when the engine
 *  refuses. A total row the engine fills with fresh defaults takes `totalsLabel` as its label. */
const toggleBoundaryRow = (
  store: SpreadsheetStore,
  history: History | null,
  wb: WorkbookHandle,
  table: TableOverlay,
  option: 'showHeader' | 'showTotal',
  totalsLabel: string,
): TableToggleResult => {
  const turningOn = !table[option];
  const fresh =
    option === 'showTotal' &&
    turningOn &&
    !engineTotals(wb, table).some((t) => t.function !== 'none' || !!t.label);
  let ok = false;
  recordTablesChange(history, store, () => {
    updateTableOverlay(store, table.id, { [option]: turningOn });
    const after = tableOverlayById(store.getState(), table.id);
    ok = after?.[option] === turningOn;
    const totals = after?.totals;
    const label = totals?.[0]?.label;
    if (ok && fresh && totals && label && label !== totalsLabel) {
      updateTableOverlay(store, table.id, {
        totals: [{ ...(totals[0] as TableColumnTotals), label: totalsLabel }, ...totals.slice(1)],
      });
    }
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
