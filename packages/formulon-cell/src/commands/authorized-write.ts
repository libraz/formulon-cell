import type { SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { interactionControllerFor } from './interaction-controller.js';
import {
  type ChangeBatchResult,
  type InteractionOrigin,
  type OperationIntent,
  type PermissionDecision,
  refusedByDecision,
} from './interaction-policy.js';

/** The intents of one logical write for an origin: the action's own origin
 *  for the apply check, `'undo'` and `'redo'` for history replay. */
export type AuthorizedIntentFactory = (origin: InteractionOrigin) => readonly OperationIntent[];

export type RefusedDecision = Extract<PermissionDecision, { allowed: false }>;

export type AuthorizedWriteResult<T> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly decision: RefusedDecision };

export interface AuthorizedWriteRequest<T> {
  /** Store whose registered interaction controller decides the intents. */
  readonly store: SpreadsheetStore;
  readonly history: History;
  readonly intents: AuthorizedIntentFactory;
  /** Origin of the action being applied. */
  readonly origin: InteractionOrigin;
  readonly write: () => T;
}

/** The first refusal among the apply, undo and redo intents of a write, so a
 *  write that could not be undone or redone is refused before it runs. Without
 *  a registered controller every write is allowed. */
export function authorizeWrite(
  store: SpreadsheetStore,
  intents: AuthorizedIntentFactory,
  origin: InteractionOrigin,
): PermissionDecision {
  const controller = interactionControllerFor(store);
  if (!controller) return { allowed: true };
  for (const intent of [...intents(origin), ...intents('undo'), ...intents('redo')]) {
    const decision = controller.canExecute(intent);
    if (!decision.allowed) return decision;
  }
  return { allowed: true };
}

/**
 * Run `write` as one history step whose undo and redo replay under the same
 * intents. The write is refused up front when any apply, undo or redo intent is
 * denied; when it throws, everything it recorded is aborted and the error is
 * rethrown. Inside an already open transaction the write joins it and the
 * outer transaction owns the replay authorization.
 */
export function authorizedWrite<T>(request: AuthorizedWriteRequest<T>): AuthorizedWriteResult<T> {
  const { store, history, intents, origin, write } = request;
  const decision = authorizeWrite(store, intents, origin);
  if (!decision.allowed) return { ok: false, decision };
  const token = history.begin(
    history.isInTransaction()
      ? undefined
      : { replayAuthorization: { undo: intents('undo'), redo: intents('redo') } },
  );
  let value: T;
  try {
    value = write();
  } catch (error) {
    try {
      history.abort(token);
    } catch (rollback) {
      throw new AggregateError([error, rollback], 'Authorized write rollback failed.');
    }
    throw error;
  }
  history.end(token);
  return { ok: true, value };
}

/** {@link authorizedWrite} of a batch command: the write's own result, or the refusal as a batch
 *  result naming the decision that refused it. */
export function authorizedBatch(
  request: AuthorizedWriteRequest<ChangeBatchResult>,
): ChangeBatchResult {
  const result = authorizedWrite(request);
  return result.ok ? result.value : refusedByDecision(result.decision);
}
