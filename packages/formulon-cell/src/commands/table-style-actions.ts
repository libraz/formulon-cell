import { type TableOverlay, tableStyleOptionsById, updateTableOverlay } from './format-as-table.js';
import { recordTablesChange } from './slice-history.js';
import { type TableActionContext, tableWrite } from './table-actions.js';

/** Gives `table` the style `styleId` as one undo step. The table keeps its own style options, as
 *  a gallery pick does. False when the change is refused. */
export function applyTableStyle(
  ctx: TableActionContext,
  table: TableOverlay,
  styleId: string,
): boolean {
  const {
    banded: _banded,
    firstCol: _firstCol,
    ...look
  } = tableStyleOptionsById(ctx.store.getState(), styleId);
  return tableWrite(ctx, table.range, false, () => {
    recordTablesChange(ctx.history, ctx.store, () => {
      updateTableOverlay(ctx.store, table.id, look);
    });
    return true;
  });
}

/** Clear in the table style gallery: `table` keeps its range and options and loses its style, as
 *  one undo step. False when the change is refused. */
export function clearTableStyle(ctx: TableActionContext, table: TableOverlay): boolean {
  return tableWrite(ctx, table.range, false, () => {
    recordTablesChange(ctx.history, ctx.store, () => {
      updateTableOverlay(ctx.store, table.id, { styleName: '' });
    });
    return true;
  });
}
