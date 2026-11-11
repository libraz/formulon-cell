import type { Addr } from '../engine/types.js';
import type { SpreadsheetStore } from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';

/** Where a note or comment change runs and which entry point asked for it. */
export interface CommentActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/**
 * Run a note or threaded-comment change of `cells` as one authorized `comment` step whose undo
 * and redo replay under the same check. Comments have no Protect Sheet option: on a protected
 * sheet each cell's lock decides it. `write` records its own undo entries; `refused` is returned
 * when the change is not allowed.
 */
export function commentWrite<T>(
  ctx: CommentActionContext,
  cells: readonly Addr[],
  refused: T,
  write: () => T,
): T {
  if (cells.length === 0) return refused;
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: (origin) => [
      {
        operation: 'comment',
        origin,
        commandId: ctx.commandId,
        effects: [{ kind: 'cells', cells }],
      },
    ],
    origin: ctx.origin,
    write,
  });
  return result.ok ? result.value : refused;
}
