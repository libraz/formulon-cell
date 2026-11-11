import type { Addr } from '../engine/types.js';
import type { SpreadsheetStore } from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import { History } from './history.js';
import type { InteractionOrigin, OperationIntent } from './interaction-policy.js';
import { asThreadIntent } from './thread-intent.js';

/** Where a note or comment change runs and which entry point asked for it. */
export interface CommentActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/**
 * Run a note change of `cells` as one authorized `comment` step whose undo and redo replay under
 * the same check. Notes have no Protect Sheet option: on a protected sheet each cell's lock
 * decides it. `write` records its own undo entries; `refused` is returned
 * when the change is not allowed.
 */
export function commentWrite<T>(
  ctx: CommentActionContext,
  cells: readonly Addr[],
  refused: T,
  write: () => T,
): T {
  return commentStep(ctx, cells, refused, write, false);
}

/**
 * Run a threaded-comment change of `cells` (add, reply, edit, resolve, delete) as one authorized
 * `comment` step. Sheet protection never refuses a thread, whatever the cell's lock or Edit
 * objects; a policy decides it as any comment change.
 */
export function threadWrite<T>(
  ctx: CommentActionContext,
  cells: readonly Addr[],
  refused: T,
  write: () => T,
): T {
  return commentStep(ctx, cells, refused, write, true);
}

function commentStep<T>(
  ctx: CommentActionContext,
  cells: readonly Addr[],
  refused: T,
  write: () => T,
  thread: boolean,
): T {
  if (cells.length === 0) return refused;
  const intent = (origin: InteractionOrigin): OperationIntent => ({
    operation: 'comment',
    origin,
    commandId: ctx.commandId,
    effects: [{ kind: 'cells', cells }],
  });
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: (origin) => [thread ? asThreadIntent(intent(origin)) : intent(origin)],
    origin: ctx.origin,
    write,
  });
  return result.ok ? result.value : refused;
}
