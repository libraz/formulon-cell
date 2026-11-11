import type { OperationIntent } from './interaction-policy.js';

// A symbol-keyed flag survives the `{ ...intent }` copies the authorizer makes per effect.
const THREAD = Symbol('threadedComment');

/** Mark `intent` as a threaded-comment change: sheet protection never decides one, whatever the
 *  cell's lock or Edit objects; a policy still does, as for any comment change. */
export const asThreadIntent = <T extends OperationIntent>(intent: T): T =>
  Object.assign(intent, { [THREAD]: true });

/** Whether `intent` is a threaded-comment change (see {@link asThreadIntent}). */
export const isThreadIntent = (intent: OperationIntent): boolean =>
  (intent as OperationIntent & { [THREAD]?: true })[THREAD] === true;
