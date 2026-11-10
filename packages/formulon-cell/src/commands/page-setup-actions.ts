import type { SpreadsheetStore } from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import { History } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';

/** Where a page-setup change runs and which entry point asked for it. */
export interface PageSetupActionContext {
  readonly store: SpreadsheetStore;
  readonly history?: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/**
 * Run a page-setup change of `sheet` as one authorized step: margins,
 * orientation, size, scale, print area and titles, header/footer, print
 * gridlines/headings, page breaks and the background image. The write records
 * its own undo entry as before; `refused` is returned when it is not allowed.
 */
export function pageSetupWrite<T>(
  ctx: PageSetupActionContext,
  sheet: number,
  refused: T,
  write: () => T,
): T {
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: (origin) => [
      {
        operation: 'pageSetup',
        origin,
        commandId: ctx.commandId,
        effects: [{ kind: 'sheet', sheet }],
      },
    ],
    origin: ctx.origin,
    write,
  });
  return result.ok ? result.value : refused;
}
