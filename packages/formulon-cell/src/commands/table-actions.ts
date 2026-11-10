import type { Range } from '../engine/types.js';
import { rangesIntersect } from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  authorizedWrite,
  authorizeWrite,
} from './authorized-write.js';
import { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';

/** Where a table change runs and which entry point asked for it. */
export interface TableActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

const tableIntents =
  (ctx: TableActionContext, range: Range): AuthorizedIntentFactory =>
  (origin) => [
    { operation: 'table', origin, commandId: ctx.commandId, effects: [{ kind: 'range', range }] },
  ];

/** Whether `ctx` may change the table over `range`, checked before a gesture starts. */
export const tableChangeAllowed = (ctx: TableActionContext, range: Range): boolean =>
  authorizeWrite(ctx.store, tableIntents(ctx, range), ctx.origin).allowed;

/**
 * Run a change of the table over `range` (or a new table there) as one authorized `table` step.
 * On a protected sheet the command alone decides it, whatever the lock of the cells. `write`
 * records its own undo entries; `refused` is returned when the change is not allowed.
 */
export function tableWrite<T>(
  ctx: TableActionContext,
  range: Range,
  refused: T,
  write: () => T,
): T {
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: tableIntents(ctx, range),
    origin: ctx.origin,
    write,
  });
  return result.ok ? result.value : refused;
}

/** Whether `range` (in either corner order) overlaps an existing table: a new table, or a Data
 *  Table, cannot go there. */
export const overlapsTable = (state: State, range: Range): boolean => {
  const normal = {
    sheet: range.sheet,
    r0: Math.min(range.r0, range.r1),
    c0: Math.min(range.c0, range.c1),
    r1: Math.max(range.r0, range.r1),
    c1: Math.max(range.c0, range.c1),
  };
  return state.tables.tables.some((t) => rangesIntersect(t.range, normal));
};

/** Whether the selection overlaps an existing table, which refuses creating a table there. */
export const selectionOverlapsTable = (state: State): boolean =>
  [state.selection.range, ...(state.selection.extraRanges ?? [])].some((range) =>
    overlapsTable(state, range),
  );
