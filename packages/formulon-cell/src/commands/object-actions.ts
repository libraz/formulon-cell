import type { SpreadsheetStore } from '../store/store.js';
import {
  type AuthorizedIntentFactory,
  authorizedWriteOr,
  authorizeWrite,
} from './authorized-write.js';
import type { History } from './history.js';
import { interactionControllerFor } from './interaction-controller.js';
import type { InteractionOrigin, OperationEffect } from './interaction-policy.js';
import { sheetPermits } from './protection.js';

/** Where a drawing-object change runs and which entry point asked for it. */
export interface ObjectActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The intents of an object change on `sheet`: the sheet as a whole (Edit objects decides it on a
 *  protected sheet), plus any cells it writes. */
export const objectIntents =
  (
    ctx: Pick<ObjectActionContext, 'commandId'>,
    sheet: number,
    cells: readonly OperationEffect[] = [],
  ): AuthorizedIntentFactory =>
  (origin) => [
    {
      operation: 'object',
      origin,
      commandId: ctx.commandId,
      effects: [{ kind: 'sheet', sheet }, ...cells],
    },
  ];

/** Whether `ctx` may change the drawing objects of `sheet`, checked before a gesture or dialog.
 *  Without a registered controller, sheet protection still decides it. */
export function objectChangeAllowed(
  ctx: Pick<ObjectActionContext, 'store' | 'origin' | 'commandId'>,
  sheet: number,
  cells: readonly OperationEffect[] = [],
): boolean {
  if (!interactionControllerFor(ctx.store))
    return sheetPermits(ctx.store.getState(), sheet, 'objects');
  return authorizeWrite(ctx.store, objectIntents(ctx, sheet, cells), ctx.origin).allowed;
}

/**
 * Run a change of the drawing objects of `sheet` (charts, pictures, shapes, ink, slicers,
 * sparklines) as one authorized `object` step whose undo and redo replay under the same check.
 * `cells` adds the cells the change writes (a sparkline's location). `write` records its own undo
 * entries; `refused` is returned when the change is not allowed.
 */
export function objectWrite<T>(
  ctx: ObjectActionContext,
  sheet: number,
  refused: T,
  write: () => T,
  cells: readonly OperationEffect[] = [],
): T {
  if (!objectChangeAllowed(ctx, sheet, cells)) return refused;
  return authorizedWriteOr(ctx, objectIntents(ctx, sheet, cells), refused, write);
}
