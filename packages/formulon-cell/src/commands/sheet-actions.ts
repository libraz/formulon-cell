import { SheetVisibility } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { type AuthorizedIntentFactory, authorizedWriteOr } from './authorized-write.js';
import type { BuiltInCommandId } from './built-in-command-ids.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import {
  addSheet,
  addSheetAfterActive,
  copySheet,
  moveSheet,
  removeSheet,
  renameSheet,
  setSheetHidden,
  setSheetTabVisibility,
} from './sheet-mutate.js';
import { recordLayoutChange } from './slice-history.js';

/** Where a sheet command runs and which entry point asked for it. */
export interface SheetActionContext {
  readonly store: SpreadsheetStore;
  readonly workbook: WorkbookHandle;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
}

type SheetCommandId = Extract<
  BuiltInCommandId,
  | 'insertSheet'
  | 'copySheet'
  | 'deleteSheet'
  | 'renameSheet'
  | 'moveSheet'
  | 'hideSheet'
  | 'unhideSheet'
  | 'sheetTabColor'
>;

/** The intent of a sheet command: a workbook-wide `sheet` operation. */
const sheetIntents =
  (commandId: SheetCommandId): AuthorizedIntentFactory =>
  (origin) => [{ operation: 'sheet', origin, commandId, effects: [{ kind: 'workbook' }] }];

const run = <T>(
  ctx: SheetActionContext,
  commandId: SheetCommandId,
  refused: T,
  write: () => T,
): T => {
  return authorizedWriteOr(ctx, sheetIntents(commandId), refused, write);
};

/** Insert a sheet, after the active one when `afterActive`; its index, or -1 when refused. */
export function insertSheetAction(ctx: SheetActionContext, afterActive = false): number {
  const { store, workbook, history } = ctx;
  return run(ctx, 'insertSheet', -1, () =>
    afterActive
      ? addSheetAfterActive(store, workbook, history)
      : addSheet(store, workbook, history),
  );
}

/** Copy the sheet at `source` before the sheet at `before` (the sheet count appends) and
 *  switch to the copy; its index, or -1 when refused. */
export function copySheetAction(ctx: SheetActionContext, source: number, before: number): number {
  return run(ctx, 'copySheet', -1, () =>
    copySheet(ctx.store, ctx.workbook, source, before, ctx.history),
  );
}

/** Delete the sheet at `idx`. */
export function deleteSheetAction(ctx: SheetActionContext, idx: number): boolean {
  return run(ctx, 'deleteSheet', false, () => removeSheet(ctx.store, ctx.workbook, idx));
}

/** Rename the sheet at `idx`. */
export function renameSheetAction(ctx: SheetActionContext, idx: number, name: string): boolean {
  return run(ctx, 'renameSheet', false, () =>
    renameSheet(ctx.workbook, idx, name, ctx.store, ctx.history),
  );
}

/** Move the sheet at `from` to `to`. */
export function moveSheetAction(ctx: SheetActionContext, from: number, to: number): boolean {
  return run(ctx, 'moveSheet', false, () =>
    moveSheet(ctx.store, ctx.workbook, from, to, ctx.history),
  );
}

/** Hide, very-hide or unhide the sheet at `idx`. */
export function setSheetVisibilityAction(
  ctx: SheetActionContext,
  idx: number,
  visibility: SheetVisibility,
): boolean {
  const commandId = visibility === SheetVisibility.Visible ? 'unhideSheet' : 'hideSheet';
  return run(ctx, commandId, false, () =>
    setSheetTabVisibility(ctx.store, ctx.workbook, ctx.history, idx, visibility),
  );
}

/** Two-state form of {@link setSheetVisibilityAction}. */
export function setSheetHiddenAction(
  ctx: SheetActionContext,
  idx: number,
  hidden: boolean,
): boolean {
  return run(ctx, hidden ? 'hideSheet' : 'unhideSheet', false, () =>
    setSheetHidden(ctx.store, ctx.workbook, ctx.history, idx, hidden),
  );
}

/** Set or clear the tab color of the sheet at `idx`. */
export function setSheetTabColorAction(
  ctx: SheetActionContext,
  idx: number,
  color: string | null,
): boolean {
  return run(ctx, 'sheetTabColor', false, () => {
    recordLayoutChange(ctx.history, ctx.store, () => {
      mutators.setSheetTabColor(ctx.store, idx, color);
    });
    return true;
  });
}
