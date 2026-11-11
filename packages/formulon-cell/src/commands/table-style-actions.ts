import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators } from '../store/store.js';
import { clearVisualFormatInRange, withSelectionFormatOrigin } from './format.js';
import {
  applyPivotTableStyleById,
  type CustomTableStyle,
  customPivotTableStyleById,
  customPivotTableStyleId,
  customTableStyleId,
  DEFAULT_TABLE_COLOR,
  defaultTableStyleId,
  overlayStyleName,
  pivotTableStyleAssignment,
  type TableOverlay,
  tableStyleOptionsById,
  updateTableOverlay,
} from './format-as-table.js';
import { updatePivotStyleInfo } from './pivot-style-info.js';
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

/** The overlay look patch of a style pick; a custom style leaves the family and hue at the defaults. */
function overlayLook(
  options: ReturnType<typeof tableStyleOptionsById>,
): Pick<TableOverlay, 'style' | 'color' | 'styleName'> {
  return {
    style: options.style ?? 'medium',
    color: options.color ?? DEFAULT_TABLE_COLOR,
    styleName: options.styleName,
  };
}

function restyle(ctx: TableActionContext, table: TableOverlay, styleId: string): boolean {
  const look = overlayLook(tableStyleOptionsById(ctx.store.getState(), styleId));
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

/** The PivotTable style a PivotTable takes when the workbook names no default. */
export const FALLBACK_PIVOT_STYLE = 'PivotStyleLight16';

type DeleteStyleWorkbook = Pick<
  WorkbookHandle,
  | 'getTableStyleDefaults'
  | 'setTableStyleDefaults'
  | 'getPivotTables'
  | 'getPivotStyleInfo'
  | 'setPivotStyleInfo'
  | 'cells'
>;

/** Deletes the custom style `name`, as one undo step, as the desktop app does. A style may be
 *  flagged for tables, PivotTables or both. The users of each kind take that kind's workbook
 *  default. A deleted default resets first, to Table Style Medium 2 / PivotTable Style Light 16.
 *  False when no custom style has that name. */
export function deleteCustomStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  workbook: DeleteStyleWorkbook | null,
  name: string,
): boolean {
  const state = ctx.store.getState();
  const named = (style: CustomTableStyle): boolean =>
    style.label.toLocaleLowerCase() === name.toLocaleLowerCase();
  const forTables = (state.tables.customTableStyles ?? []).some(named);
  const forPivots = (state.tables.customPivotTableStyles ?? []).some(named);
  if (!forTables && !forPivots) return false;
  const same = (value: string): boolean => value.toLocaleLowerCase() === name.toLocaleLowerCase();
  const history = ctx.history;
  const transaction = history?.begin();
  try {
    const before = workbook?.getTableStyleDefaults() ?? null;
    const reset = {
      ...(forTables && before && same(before.tableStyle) ? { tableStyle: '' } : {}),
      ...(forPivots && before && same(before.pivotStyle) ? { pivotStyle: '' } : {}),
    };
    if (workbook && before && Object.keys(reset).length > 0) {
      workbook.setTableStyleDefaults(reset);
      const restore = { tableStyle: before.tableStyle, pivotStyle: before.pivotStyle };
      if (history && !history.isReplaying()) {
        history.push({
          undo: () => workbook.setTableStyleDefaults(restore),
          redo: () => workbook.setTableStyleDefaults(reset),
        });
      }
    }
    const pivotUsers =
      forPivots && workbook
        ? workbook
            .getPivotTables()
            .filter((p) => same(workbook.getPivotStyleInfo(p.sheetIndex, p.pivotIndex)?.name ?? ''))
        : [];
    recordTablesChange(history, ctx.store, () => {
      mutators.removeCustomStylesNamed(ctx.store, name);
      if (!forTables) return;
      const fallback = defaultTableStyleId(ctx.store.getState(), workbook);
      const look = overlayLook(tableStyleOptionsById(ctx.store.getState(), fallback));
      for (const table of ctx.store.getState().tables.tables) {
        if (same(overlayStyleName(table))) updateTableOverlay(ctx.store, table.id, look);
      }
    });
    if (workbook) {
      const pivotDefault = workbook.getTableStyleDefaults()?.pivotStyle || FALLBACK_PIVOT_STYLE;
      for (const pivot of pivotUsers) {
        updatePivotStyleInfo(
          { store: ctx.store, workbook: workbook as WorkbookHandle, history },
          pivot,
          { name: pivotDefault },
        );
      }
    }
  } finally {
    if (history && transaction) history.end(transaction);
  }
  return true;
}

/** Replaces the custom style `name` with `next`, as Modify does, as one undo step. Every kind
 *  the style is flagged for takes the change; its tables and PivotTables follow under the new
 *  name. False when no custom style has that name. */
export function modifyCustomStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  name: string,
  next: Omit<CustomTableStyle, 'id'>,
): boolean {
  const state = ctx.store.getState();
  const named = (style: CustomTableStyle): boolean =>
    style.label.toLocaleLowerCase() === name.toLocaleLowerCase();
  const table = (state.tables.customTableStyles ?? []).find(named);
  const pivot = (state.tables.customPivotTableStyles ?? []).find(named);
  const label = next.label.trim();
  if ((!table && !pivot) || !label) return false;
  recordTablesChange(ctx.history, ctx.store, () => {
    const assigned = (ctx.store.getState().tables.pivotTableStyles ?? []).filter(
      (a) => a.styleId === pivot?.id,
    );
    mutators.removeCustomStylesNamed(ctx.store, name);
    if (table) {
      const id = customTableStyleId(label);
      mutators.upsertCustomTableStyle(ctx.store, { ...next, id, label });
      const look = overlayLook(tableStyleOptionsById(ctx.store.getState(), id));
      for (const overlay of ctx.store.getState().tables.tables) {
        if (overlayStyleName(overlay).toLocaleLowerCase() === name.toLocaleLowerCase()) {
          updateTableOverlay(ctx.store, overlay.id, look);
        }
      }
    }
    if (pivot) {
      const id = customPivotTableStyleId(label);
      mutators.upsertCustomPivotTableStyle(ctx.store, { ...next, id, label });
      for (const a of assigned) mutators.upsertPivotTableStyle(ctx.store, { ...a, styleId: id });
    }
  });
  return true;
}

/** Gives the PivotTable `pivot` the style `styleId` as one undo step: a custom style's id, a
 *  built-in name, or '' for None. */
export function applyPivotStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  workbook: WorkbookHandle,
  pivot: { sheetIndex: number; pivotIndex: number },
  styleId: string,
): boolean {
  const state = ctx.store.getState();
  const at = { sheetIndex: pivot.sheetIndex, pivotIndex: pivot.pivotIndex };
  if (customPivotTableStyleById(state, styleId)) {
    return applyPivotTableStyleById(ctx.store, ctx.history, at, styleId);
  }
  const history = ctx.history;
  const transaction = history?.begin();
  try {
    // A built-in style replaces a custom assignment, which would otherwise win on the next sync.
    if (pivotTableStyleAssignment(state, at.sheetIndex, at.pivotIndex)) {
      recordTablesChange(history, ctx.store, () => mutators.removePivotTableStyle(ctx.store, at));
    }
    return updatePivotStyleInfo({ store: ctx.store, workbook, history }, at, { name: styleId });
  } finally {
    if (history && transaction) history.end(transaction);
  }
}

/** Makes the style `name` (a built-in name or a custom style's label) its kind's workbook default,
 *  as Set As Default does. False when the engine keeps no defaults. */
export function setStyleDefault(
  workbook: Pick<WorkbookHandle, 'setTableStyleDefaults'> | null,
  kind: 'table' | 'pivot',
  name: string,
): boolean {
  if (!workbook || name === '') return false;
  return workbook.setTableStyleDefaults(
    kind === 'table' ? { tableStyle: name } : { pivotStyle: name },
  );
}
