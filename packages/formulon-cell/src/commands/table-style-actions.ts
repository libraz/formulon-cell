import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators } from '../store/store.js';
import { clearVisualFormatInRange, withSelectionFormatOrigin } from './format.js';
import {
  applyPivotTableStyleById,
  type CustomTableStyle,
  createPivotTableStyleFromActivePivot,
  createTableStyleFromActiveTable,
  currentPivotStyleId,
  currentTableStyleId,
  customPivotTableStyleById,
  customPivotTableStyleId,
  customStyleNameError,
  customTableStyleId,
  DEFAULT_TABLE_COLOR,
  defaultTableStyleId,
  overlayStyleName,
  pivotTableStyleAssignment,
  sameStyleName,
  type TableOverlay,
  tableOverlayAt,
  tableStyleOptionsById,
  updateTableOverlay,
} from './format-as-table.js';
import type { History } from './history.js';
import { updatePivotStyleInfo } from './pivot-style-info.js';
import { anySheetProtected } from './protection.js';
import { recordFormatChange, recordTablesChange } from './slice-history.js';
import { type TableActionContext, tableChangeAllowed, tableWrite } from './table-actions.js';

/** A table or a PivotTable: the two kinds of object a table style gallery styles. */
export type TableStyleKind = 'table' | 'pivot';

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
  let applied = false;
  try {
    clearTableRangeFormatting(ctx, table.range);
    applied = restyle(ctx, table, styleId);
    return applied;
  } finally {
    if (history && transaction) {
      // A refused restyle leaves the cells' formatting as it was.
      if (applied) history.end(transaction);
      else history.abort(transaction);
    }
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
  options: NonNullable<ReturnType<typeof tableStyleOptionsById>>,
): Pick<TableOverlay, 'style' | 'color' | 'styleName'> {
  return {
    style: options.style ?? 'medium',
    color: options.color ?? DEFAULT_TABLE_COLOR,
    styleName: options.styleName,
  };
}

/** The overlay look of the style `styleId`; null for an id that names no style. */
function styleLook(
  state: Parameters<typeof tableStyleOptionsById>[0],
  styleId: string,
): Pick<TableOverlay, 'style' | 'color' | 'styleName'> | null {
  const options = tableStyleOptionsById(state, styleId);
  return options ? overlayLook(options) : null;
}

function restyle(ctx: TableActionContext, table: TableOverlay, styleId: string): boolean {
  const look = styleLook(ctx.store.getState(), styleId);
  if (!look) return false;
  return tableWrite(ctx, table.range, false, () => {
    let applied = false;
    recordTablesChange(ctx.history, ctx.store, () => {
      applied = updateTableOverlay(ctx.store, table.id, look) !== null;
    });
    return applied;
  });
}

/** Clear in the table style gallery: `table` keeps its range and options and loses its style, as
 *  one undo step. False when the change is refused. */
export function clearTableStyle(ctx: TableActionContext, table: TableOverlay): boolean {
  return tableWrite(ctx, table.range, false, () => {
    let cleared = false;
    recordTablesChange(ctx.history, ctx.store, () => {
      cleared = updateTableOverlay(ctx.store, table.id, { styleName: '' }) !== null;
    });
    return cleared;
  });
}

/** The PivotTable style a PivotTable takes when the workbook names no default. */
export const FALLBACK_PIVOT_STYLE = 'PivotStyleLight16';

type StyleDefaultsWorkbook = Pick<
  WorkbookHandle,
  'getTableStyleDefaults' | 'setTableStyleDefaults'
>;

type DeleteStyleWorkbook = StyleDefaultsWorkbook &
  Pick<WorkbookHandle, 'getPivotTables' | 'getPivotStyleInfo' | 'setPivotStyleInfo' | 'cells'>;

/** Why Delete or Modify left a custom style as it was. */
export type CustomStyleRefusal = 'notFound' | 'invalidName' | 'workbookHasProtectedSheet';

/** What Delete or Modify of a custom style did. */
export type CustomStyleChange = { ok: true } | { ok: false; reason: CustomStyleRefusal };

const refused = (reason: CustomStyleRefusal): CustomStyleChange => ({ ok: false, reason });

/** Writes the workbook's style defaults that differ from what it holds and records the change on
 *  `history` (the open transaction, if any) so it undoes with the style change. An empty name
 *  clears a default. False when the engine keeps no defaults. */
function recordStyleDefaults(
  history: History | null,
  workbook: StyleDefaultsWorkbook,
  change: { tableStyle?: string; pivotStyle?: string },
): boolean {
  const before = workbook.getTableStyleDefaults();
  if (!before) return false;
  const next: { tableStyle?: string; pivotStyle?: string } = {};
  const restore: { tableStyle?: string; pivotStyle?: string } = {};
  for (const key of ['tableStyle', 'pivotStyle'] as const) {
    const value = change[key];
    if (value === undefined || value === before[key]) continue;
    next[key] = value;
    restore[key] = before[key];
  }
  if (Object.keys(next).length === 0) return true;
  if (!workbook.setTableStyleDefaults(next)) return false;
  if (history && !history.isReplaying()) {
    history.push({
      undo: () => workbook.setTableStyleDefaults(restore),
      redo: () => workbook.setTableStyleDefaults(next),
    });
  }
  return true;
}

/** Runs `fn` as one undo step when there is a history. */
const inTransaction = (history: History | null, fn: () => void): void => {
  if (history) history.transact(fn);
  else fn();
};

/** Deletes the custom style `name`, as one undo step, as the desktop app does. A style may be
 *  flagged for tables, PivotTables or both. The users of each kind take that kind's workbook
 *  default. A deleted default resets first, to Table Style Medium 2 / PivotTable Style Light 16.
 *  Refused when no custom style has that name, or while any sheet of the workbook is protected,
 *  whoever uses the style. */
export function deleteCustomStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  workbook: DeleteStyleWorkbook | null,
  name: string,
): CustomStyleChange {
  const state = ctx.store.getState();
  const named = (style: CustomTableStyle): boolean => sameStyleName(style.label, name);
  const forTables = (state.tables.customTableStyles ?? []).some(named);
  const forPivots = (state.tables.customPivotTableStyles ?? []).some(named);
  if (!forTables && !forPivots) return refused('notFound');
  if (anySheetProtected(state)) return refused('workbookHasProtectedSheet');
  const history = ctx.history;
  const transaction = history?.begin();
  try {
    const before = workbook?.getTableStyleDefaults() ?? null;
    if (workbook && before) {
      recordStyleDefaults(history, workbook, {
        ...(forTables && sameStyleName(before.tableStyle, name) ? { tableStyle: '' } : {}),
        ...(forPivots && sameStyleName(before.pivotStyle, name) ? { pivotStyle: '' } : {}),
      });
    }
    const pivotUsers =
      forPivots && workbook
        ? workbook
            .getPivotTables()
            .filter((p) =>
              sameStyleName(
                workbook.getPivotStyleInfo(p.sheetIndex, p.pivotIndex)?.name ?? '',
                name,
              ),
            )
        : [];
    // Re-point the PivotTables while the style still exists, so undo can point them back: the
    // engine re-points any it still finds when the style goes.
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
    recordTablesChange(history, ctx.store, () => {
      mutators.removeCustomStylesNamed(ctx.store, name);
      if (!forTables) return;
      const fallback = defaultTableStyleId(ctx.store.getState(), workbook);
      const look = styleLook(ctx.store.getState(), fallback);
      if (!look) return;
      for (const table of ctx.store.getState().tables.tables) {
        if (sameStyleName(overlayStyleName(table), name))
          updateTableOverlay(ctx.store, table.id, look);
      }
    });
  } finally {
    if (history && transaction) history.end(transaction);
  }
  return { ok: true };
}

/** How a Modify dialog's Set As Default checkbox applies to `kind`'s workbook default. */
export interface ModifyStyleDefault {
  workbook: StyleDefaultsWorkbook | null;
  kind: TableStyleKind;
  setAsDefault: boolean;
}

/** Replaces the custom style `name` with `next`, as Modify does, as one undo step. Every kind
 *  the style is flagged for takes the change; its tables and PivotTables follow under the new
 *  name, and so does a workbook default that named it. `defaults` applies the dialog's Set As
 *  Default checkbox: ticked makes the style `kind`'s default, unticked clears that default when
 *  it named the style. Refused when no custom style has that name, the new name is taken, or any
 *  sheet of the workbook is protected. */
export function modifyCustomStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  name: string,
  next: Omit<CustomTableStyle, 'id'>,
  defaults?: ModifyStyleDefault,
): CustomStyleChange {
  const state = ctx.store.getState();
  const named = (style: CustomTableStyle): boolean => sameStyleName(style.label, name);
  const table = (state.tables.customTableStyles ?? []).find(named);
  const pivot = (state.tables.customPivotTableStyles ?? []).find(named);
  const label = next.label.trim();
  if (!table && !pivot) return refused('notFound');
  if (anySheetProtected(state)) return refused('workbookHasProtectedSheet');
  if (!label || customStyleNameError(state, label, { exceptLabel: name }))
    return refused('invalidName');
  inTransaction(ctx.history, () => {
    // The engine resets a default whose style it removes: a default naming the old name moves
    // first, and a ticked one is set again once the style is back.
    const workbook = defaults?.workbook;
    const before = workbook?.getTableStyleDefaults();
    const key = defaults?.kind === 'pivot' ? 'pivotStyle' : 'tableStyle';
    if (defaults && workbook && before) {
      recordStyleDefaults(ctx.history, workbook, {
        ...(table && sameStyleName(before.tableStyle, name) ? { tableStyle: label } : {}),
        ...(pivot && sameStyleName(before.pivotStyle, name) ? { pivotStyle: label } : {}),
        ...(defaults.setAsDefault
          ? { [key]: label }
          : sameStyleName(before[key], name)
            ? { [key]: '' }
            : {}),
      });
    }
    recordTablesChange(ctx.history, ctx.store, () => {
      const assigned = (ctx.store.getState().tables.pivotTableStyles ?? []).filter(
        (a) => a.styleId === pivot?.id,
      );
      mutators.removeCustomStylesNamed(ctx.store, name);
      if (table) {
        const id = customTableStyleId(label);
        mutators.upsertCustomTableStyle(ctx.store, { ...next, id, label });
        const look = styleLook(ctx.store.getState(), id);
        for (const overlay of ctx.store.getState().tables.tables) {
          if (look && sameStyleName(overlayStyleName(overlay), name)) {
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
    if (defaults?.setAsDefault && workbook)
      recordStyleDefaults(ctx.history, workbook, { [key]: label });
  });
  return { ok: true };
}

/** Adds a custom style of `kind` from a dialog result, and makes it the workbook default of its
 *  kind when asked, as one undo step. False when the name is empty or taken. */
export function createCustomStyle(
  ctx: Omit<TableActionContext, 'origin' | 'commandId'>,
  workbook: StyleDefaultsWorkbook | null,
  kind: TableStyleKind,
  value: {
    name: string;
    elements: CustomTableStyle['elements'];
    stripeSizes: CustomTableStyle['stripeSizes'];
    setAsDefault: boolean;
  },
): boolean {
  const label = value.name.trim();
  if (!label || customStyleNameError(ctx.store.getState(), label)) return false;
  const look = { elements: value.elements, stripeSizes: value.stripeSizes };
  inTransaction(ctx.history, () => {
    if (kind === 'table') createTableStyleFromActiveTable(ctx.store, ctx.history, label, look);
    else createPivotTableStyleFromActivePivot(ctx.store, ctx.history, label, look);
    if (value.setAsDefault) setStyleDefault(workbook, kind, label, ctx.history);
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
    const applied = updatePivotStyleInfo({ store: ctx.store, workbook, history }, at, {
      name: styleId,
    });
    // A built-in style replaces a custom assignment, which would otherwise win on the next sync.
    if (applied && pivotTableStyleAssignment(state, at.sheetIndex, at.pivotIndex)) {
      recordTablesChange(history, ctx.store, () => mutators.removePivotTableStyle(ctx.store, at));
    }
    return applied;
  } finally {
    if (history && transaction) history.end(transaction);
  }
}

/** Makes the style `name` (a built-in name or a custom style's label) its kind's workbook default,
 *  as Set As Default does; with `history` and a workbook that reads its defaults, the change is
 *  recorded on it. False when the engine keeps no defaults. */
export function setStyleDefault(
  workbook: (Pick<WorkbookHandle, 'setTableStyleDefaults'> & Partial<StyleDefaultsWorkbook>) | null,
  kind: TableStyleKind,
  name: string,
  history: History | null = null,
): boolean {
  if (!workbook || name === '') return false;
  const change = kind === 'table' ? { tableStyle: name } : { pivotStyle: name };
  return history && workbook.getTableStyleDefaults
    ? recordStyleDefaults(history, workbook as StyleDefaultsWorkbook, change)
    : workbook.setTableStyleDefaults(change);
}

/** The table or PivotTable under the active cell, as one shape for both galleries. */
export interface ActiveStyleTarget {
  readonly kind: TableStyleKind;
  /** The gallery id of its current style: a custom style's id, a built-in name, or '' for none. */
  readonly currentId: string;
  /** Gives it the style `styleId` as one undo step; `clearFormatting` first clears its cells'
   *  direct formatting. False when the change is refused, leaving the cells as they were. */
  apply(styleId: string, options?: { readonly clearFormatting?: boolean }): boolean;
  /** Removes its style (Clear in the gallery). False when refused. */
  clear(): boolean;
}

/** The `kind` object under the active cell, or null when the active cell is outside one. */
export function activeStyleTarget(
  ctx: TableActionContext & { readonly workbook: WorkbookHandle },
  kind: TableStyleKind,
): ActiveStyleTarget | null {
  const state = ctx.store.getState();
  const active = state.selection.active;
  if (kind === 'table') {
    const table = tableOverlayAt(state, active.sheet, active.row, active.col);
    if (!table) return null;
    return {
      kind,
      currentId: currentTableStyleId(state, overlayStyleName(table)),
      // None is no style.
      apply: (styleId, options) =>
        styleId === ''
          ? clearTableStyle(ctx, table)
          : applyTableStyle(ctx, table, styleId, options),
      clear: () => clearTableStyle(ctx, table),
    };
  }
  const pivot = findPivotTableAtCell(ctx.workbook, active);
  if (!pivot) return null;
  const name = ctx.workbook.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex)?.name ?? '';
  const restyle = (styleId: string): boolean => applyPivotStyle(ctx, ctx.workbook, pivot, styleId);
  return {
    kind,
    currentId: currentPivotStyleId(state, name),
    apply: (styleId, options) => {
      if (!options?.clearFormatting) return restyle(styleId);
      const history = ctx.history;
      const transaction = history?.begin();
      let applied = false;
      try {
        clearTableRangeFormatting(ctx, {
          sheet: pivot.sheetIndex,
          r0: pivot.top,
          c0: pivot.left,
          r1: pivot.top + pivot.rows - 1,
          c1: pivot.left + pivot.cols - 1,
        });
        applied = restyle(styleId);
        return applied;
      } finally {
        if (history && transaction) {
          // A refused restyle leaves the cells' formatting as it was.
          if (applied) history.end(transaction);
          else history.abort(transaction);
        }
      }
    },
    clear: () => restyle(''),
  };
}
