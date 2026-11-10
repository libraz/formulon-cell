import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  type AuthorizedWriteResult,
  authorizedWrite,
} from './authorized-write.js';
import { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { recordDefinedNamesChange } from './named-ranges.js';

/** Where a defined-name change runs and which entry point asked for it. */
export interface NamedRangeWriteContext {
  readonly store: SpreadsheetStore;
  readonly workbook: WorkbookHandle;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The intent of a defined-name change: a workbook-wide `namedRange` operation. */
const namedRangeIntents =
  (commandId: string): AuthorizedIntentFactory =>
  (origin) => [{ operation: 'namedRange', origin, commandId, effects: [{ kind: 'workbook' }] }];

/** Run a defined-name change as one authorized, undoable step. */
export function namedRangeWrite<T>(
  ctx: NamedRangeWriteContext,
  write: () => T,
): AuthorizedWriteResult<T> {
  return authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: namedRangeIntents(ctx.commandId),
    origin: ctx.origin,
    write: () => recordDefinedNamesChange(ctx.history, ctx.workbook, write),
  });
}
