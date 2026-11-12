import type { Range } from '../engine/types.js';
import {
  isSingleCell,
  normalizeRange,
  rangesIntersect,
  selectionRangeList,
} from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  authorizedWriteOr,
  authorizeWrite,
} from './authorized-write.js';
import { currentRegion } from './current-region.js';
import type { History } from './history.js';
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
  return authorizedWriteOr(ctx, tableIntents(ctx, range), refused, write);
}

/** Whether `range` (in either corner order) overlaps an existing table: a new table, or a Data
 *  Table, cannot go there. */
export const overlapsTable = (state: State, range: Range): boolean => {
  const normal = normalizeRange(range);
  return state.tables.tables.some((t) => rangesIntersect(t.range, normal));
};

/** Whether the selection overlaps an existing table, which refuses creating a table there. */
export const selectionOverlapsTable = (state: State): boolean =>
  selectionRangeList(state.selection).some((range) => overlapsTable(state, range));

/** The range a new table proposes: the selection, or for a single cell its current region (the
 *  cell alone when it is empty), as Create Table detects it. */
export const tableSourceForSelection = (state: State): Range => {
  const r = state.selection.range;
  const range: Range = normalizeRange(r);
  if (!isSingleCell(range)) return range;
  return currentRegion(state, state.selection.active) ?? range;
};
