import type { Addr, Range } from '../engine/types.js';
import { writeCell } from '../engine/value.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { DataTableDefinition, DataTableInfo } from '../engine/workbook-handle-data-tables.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { overlapsTable } from './table-actions.js';

export type DataTableRequestError = 'invalid-range' | 'invalid-input' | 'rejected';

/** A Data Table request as the dialog collects it: the selection holding the
 *  formula row/column and the substitution values, plus the input cells. */
export interface DataTableRequest {
  selection: Range;
  rowInput: Addr | null;
  colInput: Addr | null;
}

interface DataTableCommandContext {
  store: SpreadsheetStore;
  workbook: WorkbookHandle;
  history: History | null;
}

/** The command every Data Table entry point runs as. */
export const DATA_TABLE_COMMAND = 'mac.data.dataTable';

/** A Data Table creation and the entry point that asked for it. */
interface DataTableCreateContext extends DataTableCommandContext {
  origin: InteractionOrigin;
}

const inside = (range: Range, addr: Addr): boolean =>
  addr.sheet === range.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

/** The result block is the selection minus its first row and first column. */
export function dataTableDefinition(
  request: DataTableRequest,
): DataTableDefinition | DataTableRequestError {
  const { selection, rowInput, colInput } = request;
  if (selection.r1 <= selection.r0 || selection.c1 <= selection.c0) return 'invalid-range';
  if (!rowInput && !colInput) return 'invalid-input';
  const range = { ...selection, r0: selection.r0 + 1, c0: selection.c0 + 1 };
  for (const input of [rowInput, colInput]) {
    if (input && (input.sheet !== selection.sheet || inside(range, input))) return 'invalid-input';
  }
  return { range, rowInput, colInput };
}

const refreshCells = (store: SpreadsheetStore, workbook: WorkbookHandle): void => {
  mutators.replaceCells(store, workbook.cells(store.getState().data.sheetIndex));
};

/** What-If Analysis → Data Table, as one authorized undo step. A protected sheet refuses it
 *  whatever its options, and so does a selection over an existing table. */
export function createDataTable(
  ctx: DataTableCreateContext,
  request: DataTableRequest,
): DataTableRequestError | null {
  const def = dataTableDefinition(request);
  if (typeof def === 'string') return def;
  if (overlapsTable(ctx.store.getState(), request.selection)) return 'rejected';
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: (origin) => [
      {
        operation: 'valueEdit',
        origin,
        commandId: DATA_TABLE_COMMAND,
        effects: [{ kind: 'range', range: def.range }],
      },
    ],
    origin: ctx.origin,
    write: () => writeDataTable(ctx, def),
  });
  return result.ok ? result.value : 'rejected';
}

const writeDataTable = (
  ctx: DataTableCommandContext,
  def: DataTableDefinition,
): DataTableRequestError | null => {
  const { store, workbook, history } = ctx;
  const replaced = [...workbook.cellsInRange(def.range)];
  if (!workbook.addDataTable(def)) return 'rejected';
  const anchor = { sheet: def.range.sheet, row: def.range.r0, col: def.range.c0 };
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => {
        workbook.removeDataTable(anchor);
        // The table overwrote whatever the result block held.
        workbook.withBatchedRecalc(() => {
          for (const cell of replaced) writeCell(workbook, cell.addr, cell.value, cell.formula);
        });
        refreshCells(store, workbook);
      },
      redo: () => {
        workbook.addDataTable(def);
        refreshCells(store, workbook);
      },
    });
  }
  refreshCells(store, workbook);
  return null;
};

/** Remove whole data tables, as Delete over their result blocks does. */
export function deleteDataTables(
  ctx: DataTableCommandContext,
  tables: readonly DataTableInfo[],
): boolean {
  const { store, workbook, history } = ctx;
  for (const table of tables) {
    const anchor = { sheet: table.range.sheet, row: table.range.r0, col: table.range.c0 };
    const def: DataTableDefinition = {
      range: table.range,
      rowInput: table.rowInput,
      colInput: table.colInput,
    };
    if (!workbook.removeDataTable(anchor)) return false;
    if (history && !history.isReplaying()) {
      history.push({
        undo: () => {
          workbook.addDataTable(def);
          refreshCells(store, workbook);
        },
        redo: () => {
          workbook.removeDataTable(anchor);
          refreshCells(store, workbook);
        },
      });
    }
  }
  refreshCells(store, workbook);
  return true;
}
