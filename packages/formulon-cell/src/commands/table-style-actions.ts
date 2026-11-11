import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators } from '../store/store.js';
import { clearVisualFormatInRange, withSelectionFormatOrigin } from './format.js';
import {
  type CustomTableStyle,
  customTableStyleById,
  customTableStyleId,
  defaultTableStyleId,
  overlayStyleName,
  type TableOverlay,
  tableStyleOptionsById,
  updateTableOverlay,
} from './format-as-table.js';
import { recordFormatChange, recordTablesChange } from './slice-history.js';
import { type TableActionContext, tableChangeAllowed, tableWrite } from './table-actions.js';

/** Gives `table` the style `styleId` as one undo step. The table keeps its own style options, as
 *  a gallery pick does; `clearFormatting` first clears the direct formatting of its cells, as
 *  Apply and Clear Formatting does. False when the change is refused. */
export function applyTableStyle(
  ctx: TableActionContext,
  table: TableOverlay,
  styleId: string,
  options: { readonly clearFormatting?: boolean } = {},
): boolean {
  if (!options.clearFormatting) return restyle(ctx, table, styleId);
  if (!tableChangeAllowed(ctx, table.range)) return false;
  const history = ctx.history;
  const transaction = history?.begin();
  try {
    clearTableRangeFormatting(ctx, table.range);
    return restyle(ctx, table, styleId);
  } finally {
    if (history && transaction) history.end(transaction);
  }
}

/** Clears the direct formatting of `range` with its own undo entry, the first half of Apply and
 *  Clear Formatting; the caller groups it with the style change. */
export function clearTableRangeFormatting(ctx: TableActionContext, range: Range): void {
  recordFormatChange(ctx.history, ctx.store, () =>
    withSelectionFormatOrigin(
      ctx.store,
      ctx.origin,
      () => clearVisualFormatInRange(ctx.store.getState(), ctx.store, range),
      ctx.commandId,
    ),
  );
}

function restyle(ctx: TableActionContext, table: TableOverlay, styleId: string): boolean {
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

/** Deletes the custom table style `styleId` as one undo step. Tables in it take the workbook
 *  default table style; that fallback is unmeasured. False for a built-in or unknown id. */
export function deleteCustomTableStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  workbook: Pick<WorkbookHandle, 'getTableStyleDefaults'> | null,
  styleId: string,
): boolean {
  const style = customTableStyleById(ctx.store.getState(), styleId);
  if (!style) return false;
  recordTablesChange(ctx.history, ctx.store, () => {
    mutators.removeCustomTableStyle(ctx.store, styleId);
    const fallback = defaultTableStyleId(ctx.store.getState(), workbook);
    const look = tableStyleOptionsById(ctx.store.getState(), fallback);
    for (const table of ctx.store.getState().tables.tables) {
      if (overlayStyleName(table) !== style.label) continue;
      updateTableOverlay(ctx.store, table.id, {
        style: look.style,
        color: look.color,
        styleName: look.styleName,
      });
    }
  });
  return true;
}

/** Replaces the custom table style `styleId` with `next`, as Modify does, as one undo step.
 *  Tables in it follow, under its new name. False for a built-in or unknown id. */
export function modifyCustomTableStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  styleId: string,
  next: Omit<CustomTableStyle, 'id'>,
): boolean {
  const before = customTableStyleById(ctx.store.getState(), styleId);
  const label = next.label.trim();
  if (!before || !label) return false;
  recordTablesChange(ctx.history, ctx.store, () => {
    mutators.removeCustomTableStyle(ctx.store, styleId);
    const id = customTableStyleId(label);
    mutators.upsertCustomTableStyle(ctx.store, { ...next, id, label });
    const look = tableStyleOptionsById(ctx.store.getState(), id);
    for (const table of ctx.store.getState().tables.tables) {
      if (overlayStyleName(table) !== before.label) continue;
      updateTableOverlay(ctx.store, table.id, {
        style: look.style,
        color: look.color,
        styleName: look.styleName,
      });
    }
  });
  return true;
}
