import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  type AuthorizedWriteResult,
  authorizedWrite,
} from './authorized-write.js';
import {
  circleInvalidValidationData,
  clearValidationCircles,
  recordValidationCirclesChange,
} from './error-indicators.js';
import type { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import {
  circleInvalidValidationDataWithEngine,
  clearValidationInRangeWithEngine,
} from './validate.js';

/** Where a Data Validation menu action runs and who asked for it. */
export interface ValidationActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History;
  readonly workbook: WorkbookHandle;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The intents of a validation write over `range`. */
export const validationIntents =
  (range: Range, commandId: string): AuthorizedIntentFactory =>
  (origin) => [{ operation: 'validation', origin, commandId, effects: [{ kind: 'range', range }] }];

const run = <T>(
  ctx: ValidationActionContext,
  range: Range,
  write: () => T,
): AuthorizedWriteResult<T> =>
  authorizedWrite({
    store: ctx.store,
    history: ctx.history,
    intents: validationIntents(range, ctx.commandId),
    origin: ctx.origin,
    write,
  });

/** Circle Invalid Data: add a circle to every cell in `range` that breaks its rule. */
export function circleInvalidData(
  ctx: ValidationActionContext,
  range: Range,
): AuthorizedWriteResult<void> {
  const { store, history, workbook } = ctx;
  return run(ctx, range, () =>
    recordValidationCirclesChange(history, store, () => {
      const nativeCount = circleInvalidValidationDataWithEngine(
        store,
        workbook,
        range.sheet,
        range,
      );
      if (nativeCount === null) {
        circleInvalidValidationData(store, range, makeRangeResolver(workbook, range.sheet));
      }
    }),
  );
}

/** Clear Validation Circles, decided over the selection the command runs on. */
export function clearValidationCirclesAction(
  ctx: ValidationActionContext,
  range: Range,
): AuthorizedWriteResult<void> {
  const { store, history } = ctx;
  return run(ctx, range, () =>
    recordValidationCirclesChange(history, store, () => clearValidationCircles(store)),
  );
}

/** Clear Validation: remove the rules in `range` and every validation circle. */
export function clearValidationRules(
  ctx: ValidationActionContext,
  range: Range,
): AuthorizedWriteResult<number> {
  const { store, history, workbook } = ctx;
  return run(ctx, range, () => {
    const cleared = clearValidationInRangeWithEngine(store, history, workbook, range);
    recordValidationCirclesChange(history, store, () => clearValidationCircles(store));
    return cleared;
  });
}
