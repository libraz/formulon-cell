import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SheetProtectionPermissions, SpreadsheetStore } from '../store/store.js';
import { type AuthorizedIntentFactory, authorizedWriteOr } from './authorized-write.js';
import type { History } from './history.js';
import type { InteractionOrigin, OperationEffect } from './interaction-policy.js';
import {
  type AllowedEditRangeOptions,
  type AllowedEditRangePatch,
  addAllowedEditRange,
  clearAllowedEditRanges,
  removeAllowedEditRange,
  setCellLocked,
  setProtectedSheet,
  setWorkbookStructureProtected,
  updateAllowedEditRange,
} from './protection.js';
import { recordFormatChange } from './slice-history.js';

/** Where a protection command runs and which entry point asked for it. */
export interface ProtectionActionContext {
  readonly store: SpreadsheetStore;
  readonly workbook?: WorkbookHandle;
  readonly history?: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

const protectionIntents =
  (commandId: string, effect: OperationEffect): AuthorizedIntentFactory =>
  (origin) => [{ operation: 'protection', origin, commandId, effects: [effect] }];

const run = <T>(
  ctx: ProtectionActionContext,
  effect: OperationEffect,
  refused: T,
  write: () => T,
): T => {
  return authorizedWriteOr(ctx, protectionIntents(ctx.commandId, effect), refused, write);
};

const sheetOf = (ctx: ProtectionActionContext, id: string): number | null =>
  ctx.store.getState().protection.allowedEditRanges.find((entry) => entry.id === id)?.range.sheet ??
  null;

/** Protect or unprotect `sheet`. Protection never refuses itself, so this runs on a protected sheet. */
export function setSheetProtectionAction(
  ctx: ProtectionActionContext,
  sheet: number,
  on: boolean,
  options: { readonly password?: string; readonly permissions?: SheetProtectionPermissions } = {},
): boolean {
  return run(ctx, { kind: 'sheet', sheet }, false, () => {
    setProtectedSheet(ctx.store, sheet, on, { ...options, workbook: ctx.workbook });
    return true;
  });
}

/** Protect or unprotect the workbook structure. */
export function setWorkbookProtectionAction(
  ctx: ProtectionActionContext,
  on: boolean,
  password?: string,
): boolean {
  return run(ctx, { kind: 'workbook' }, false, () => {
    setWorkbookStructureProtected(ctx.store, on, {
      ...(password !== undefined ? { password } : {}),
      workbook: ctx.workbook,
    });
    return true;
  });
}

/** Add an allowed edit range; its id, or null when refused. */
export function addAllowedEditRangeAction(
  ctx: ProtectionActionContext,
  range: Range,
  options: Omit<AllowedEditRangeOptions, 'workbook'> = {},
): string | null {
  return run(ctx, { kind: 'sheet', sheet: range.sheet }, null, () =>
    addAllowedEditRange(ctx.store, range, { ...options, workbook: ctx.workbook }),
  );
}

/** Change the allowed edit range `id`. */
export function updateAllowedEditRangeAction(
  ctx: ProtectionActionContext,
  id: string,
  patch: AllowedEditRangePatch,
): boolean {
  const sheet = sheetOf(ctx, id);
  if (sheet === null) return false;
  return run(ctx, { kind: 'sheet', sheet }, false, () =>
    updateAllowedEditRange(ctx.store, id, patch, { workbook: ctx.workbook }),
  );
}

/** Delete the allowed edit range `id`. */
export function removeAllowedEditRangeAction(ctx: ProtectionActionContext, id: string): boolean {
  const sheet = sheetOf(ctx, id);
  if (sheet === null) return false;
  return run(ctx, { kind: 'sheet', sheet }, false, () => {
    removeAllowedEditRange(ctx.store, id, { workbook: ctx.workbook });
    return true;
  });
}

/** Delete every allowed edit range of `sheet`. */
export function clearAllowedEditRangesAction(ctx: ProtectionActionContext, sheet: number): boolean {
  return run(ctx, { kind: 'sheet', sheet }, false, () => {
    clearAllowedEditRanges(ctx.store, sheet, { workbook: ctx.workbook });
    return true;
  });
}

/** Lock or unlock the cells of `range`, as one undoable format change. */
export function setCellsLockedAction(
  ctx: ProtectionActionContext,
  range: Range,
  locked: boolean,
): boolean {
  return run(ctx, { kind: 'range', range }, false, () => {
    recordFormatChange(ctx.history ?? null, ctx.store, () =>
      setCellLocked(ctx.store, range, locked),
    );
    return true;
  });
}
