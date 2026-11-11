import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  authorizedWrite,
  authorizeWrite,
} from './authorized-write.js';
import { recordFilterChange } from './filter.js';
import { History } from './history.js';
import type { InteractionOrigin, OperationEffect } from './interaction-policy.js';

/** Where a filter change runs and which entry point asked for it. */
export interface FilterActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly workbook?: WorkbookHandle;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

const run = <T>(
  ctx: FilterActionContext,
  intents: AuthorizedIntentFactory,
  refused: T,
  write: () => T,
): T => {
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents,
    origin: ctx.origin,
    write: () => recordFilterChange(ctx.history, ctx.store, write, ctx.workbook),
  });
  return result.ok ? result.value : refused;
};

const filterIntents =
  (commandId: string, effects: readonly OperationEffect[]): AuthorizedIntentFactory =>
  (origin) => [{ operation: 'filter', origin, commandId, effects }];

/**
 * Use the filter on `range`: apply criteria from the dropdown, clear them,
 * reapply, filter by a cell's value, or drive it from a slicer. Decided on the
 * filtered range, whose rows it hides or shows, so a protected sheet needs its
 * Use AutoFilter option.
 */
export function filterUseWrite<T>(
  ctx: FilterActionContext,
  range: Range,
  refused: T,
  write: () => T,
): T {
  return run(ctx, filterIntents(ctx.commandId, [{ kind: 'range', range }]), refused, write);
}

/** Whether `ctx` may use the filter on `range`, checked before its dropdown opens: on a protected
 *  sheet without Use AutoFilter the button does nothing. */
export const filterUseAllowed = (
  ctx: Pick<FilterActionContext, 'store' | 'origin' | 'commandId'>,
  range: Range,
): boolean =>
  authorizeWrite(ctx.store, filterIntents(ctx.commandId, [{ kind: 'range', range }]), ctx.origin)
    .allowed;

/**
 * Turn AutoFilter on or off on `sheet`. Decided for the sheet as a whole: no
 * Protect Sheet option enables it on a protected sheet.
 */
export function filterDefinitionWrite<T>(
  ctx: FilterActionContext,
  sheet: number,
  refused: T,
  write: () => T,
): T {
  return run(ctx, filterIntents(ctx.commandId, [{ kind: 'sheet', sheet }]), refused, write);
}

/**
 * Advanced Filter with its result copied to `output`: a filter use on
 * `listRange` and a value write over the cells the copy may cover.
 */
export function advancedFilterCopyWrite<T>(
  ctx: FilterActionContext,
  listRange: Range,
  output: Range,
  refused: T,
  write: () => T,
): T {
  return run(
    ctx,
    (origin) => [
      ...filterIntents(ctx.commandId, [{ kind: 'range', range: listRange }])(origin),
      {
        operation: 'valueEdit',
        origin,
        commandId: ctx.commandId,
        effects: [{ kind: 'range', range: output }],
      },
    ],
    refused,
    write,
  );
}
