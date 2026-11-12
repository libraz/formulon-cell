import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { authorizedWriteOr, authorizeWrite } from './authorized-write.js';
import {
  type DeleteCellsDirection,
  deleteCells,
  type InsertCellsDirection,
  insertCells,
  insertTableColumns,
  insertTableRows,
  wholeLineEdit,
} from './cell-shift.js';
import { tableForCell } from './format-as-table.js';
import type { History } from './history.js';
import { cellShiftBand, columnShiftBand, rowShiftBand } from './interaction-effects.js';
import type {
  InteractionOperation,
  InteractionOrigin,
  OperationEffect,
} from './interaction-policy.js';
import { deleteCols, deleteVisibleRows, insertCols, insertRows } from './structure.js';

/** Where a row or column change runs and which entry point asked for it. */
export interface RowColActionContext {
  readonly store: SpreadsheetStore;
  readonly workbook: WorkbookHandle;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The axis a row or column layout change acts on. */
export type LayoutAxis = 'row' | 'col';

/** Run `write` as one authorized step of `operation` over `effects`; `refused` when refused. */
export function rowColWrite<T>(
  ctx: Pick<RowColActionContext, 'store' | 'history' | 'origin' | 'commandId'>,
  operation: InteractionOperation,
  effects: readonly OperationEffect[],
  refused: T,
  write: () => T,
): T {
  return authorizedWriteOr(
    ctx,
    (origin) => [{ operation, origin, commandId: ctx.commandId, effects }],
    refused,
    write,
  );
}

/**
 * Size, hide, show or expand/collapse rows (`row`) or columns (`col`) of the active sheet as one
 * authorized step: decided for the sheet, by its Format rows/columns option when it is protected.
 */
export function layoutWrite<T>(
  ctx: Pick<RowColActionContext, 'store' | 'history' | 'origin' | 'commandId'>,
  axis: LayoutAxis,
  refused: T,
  write: () => T,
): T {
  const sheet = ctx.store.getState().data.sheetIndex;
  const operation = axis === 'row' ? 'resizeRows' : 'resizeColumns';
  return rowColWrite(ctx, operation, [{ kind: 'sheet', sheet }], refused, write);
}

/** Whether `ctx` may change the row (column) layout of the active sheet, checked before a prompt
 *  or a gesture starts. */
export function layoutAllowed(
  ctx: Pick<RowColActionContext, 'store' | 'origin' | 'commandId'>,
  axis: LayoutAxis,
): boolean {
  const sheet = ctx.store.getState().data.sheetIndex;
  const operation = axis === 'row' ? 'resizeRows' : 'resizeColumns';
  return authorizeWrite(
    ctx.store,
    (origin) => [
      { operation, origin, commandId: ctx.commandId, effects: [{ kind: 'sheet', sheet }] },
    ],
    ctx.origin,
  ).allowed;
}

const activeSheet = (ctx: RowColActionContext): number => ctx.store.getState().data.sheetIndex;

/** Insert `count` sheet rows at `at`. */
export const insertRowsAction = (ctx: RowColActionContext, at: number, count = 1): boolean =>
  rowColWrite(
    ctx,
    'insertRows',
    [{ kind: 'range', range: rowShiftBand(activeSheet(ctx), at) }],
    false,
    () => insertRows(ctx.store, ctx.workbook, ctx.history, at, count),
  );

/** Delete the visible rows of `count` sheet rows from `at`. */
export const deleteRowsAction = (ctx: RowColActionContext, at: number, count = 1): boolean =>
  rowColWrite(
    ctx,
    'deleteRows',
    [{ kind: 'range', range: rowShiftBand(activeSheet(ctx), at) }],
    false,
    () => deleteVisibleRows(ctx.store, ctx.workbook, ctx.history, at, count),
  );

/** Insert `count` sheet columns at `at`. */
export const insertColumnsAction = (ctx: RowColActionContext, at: number, count = 1): boolean =>
  rowColWrite(
    ctx,
    'insertColumns',
    [{ kind: 'range', range: columnShiftBand(activeSheet(ctx), at) }],
    false,
    () => insertCols(ctx.store, ctx.workbook, ctx.history, at, count),
  );

/** Delete `count` sheet columns from `at`. */
export const deleteColumnsAction = (ctx: RowColActionContext, at: number, count = 1): boolean =>
  rowColWrite(
    ctx,
    'deleteColumns',
    [{ kind: 'range', range: columnShiftBand(activeSheet(ctx), at) }],
    false,
    () => deleteCols(ctx.store, ctx.workbook, ctx.history, at, count),
  );

/** Insert cells over `range`, shifting the cells below down or those to the right across; whole
 *  rows or columns are inserted as such, whatever the direction. */
export function insertCellsAction(
  ctx: RowColActionContext,
  range: Range,
  direction: InsertCellsDirection,
): boolean {
  const lines = wholeLineEdit(ctx.store, range);
  if (lines?.axis === 'row') return insertRowsAction(ctx, lines.at, lines.count);
  if (lines) return insertColumnsAction(ctx, lines.at, lines.count);
  return rowColWrite(
    ctx,
    direction === 'down' ? 'insertRows' : 'insertColumns',
    [{ kind: 'range', range: cellShiftBand(range, direction === 'down') }],
    false,
    () => insertCells(ctx.store, ctx.workbook, ctx.history, range, direction),
  );
}

/** Delete the cells of `range`, shifting the cells below up or those to the right across; whole
 *  rows or columns are deleted as such, whatever the direction. */
export function deleteCellsAction(
  ctx: RowColActionContext,
  range: Range,
  direction: DeleteCellsDirection,
): boolean {
  const lines = wholeLineEdit(ctx.store, range);
  if (lines?.axis === 'row') return deleteRowsAction(ctx, lines.at, lines.count);
  if (lines) return deleteColumnsAction(ctx, lines.at, lines.count);
  return rowColWrite(
    ctx,
    direction === 'up' ? 'deleteRows' : 'deleteColumns',
    [{ kind: 'range', range: cellShiftBand(range, direction === 'up') }],
    false,
    () => deleteCells(ctx.store, ctx.workbook, ctx.history, range, direction),
  );
}

/** Table Rows Above: rows inserted within the table's columns, shifting what lies below. */
export function insertTableRowsAction(ctx: RowColActionContext, range: Range): boolean {
  const table = tableForCell(ctx.store.getState().tables.tables, range.sheet, range.r0, range.c0);
  if (!table) return false;
  const band = cellShiftBand({ ...range, c0: table.range.c0, c1: table.range.c1 }, true);
  return rowColWrite(ctx, 'insertRows', [{ kind: 'range', range: band }], false, () =>
    insertTableRows(ctx.store, ctx.workbook, ctx.history, range),
  );
}

/** Table Columns to the Left: columns inserted within the table's rows, shifting what lies right. */
export function insertTableColumnsAction(ctx: RowColActionContext, range: Range): boolean {
  const table = tableForCell(ctx.store.getState().tables.tables, range.sheet, range.r0, range.c0);
  if (!table) return false;
  const band = cellShiftBand({ ...range, r0: table.range.r0, r1: table.range.r1 }, false);
  return rowColWrite(ctx, 'insertColumns', [{ kind: 'range', range: band }], false, () =>
    insertTableColumns(ctx.store, ctx.workbook, ctx.history, range),
  );
}
