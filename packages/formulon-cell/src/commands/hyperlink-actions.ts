import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import { History } from './history.js';
import { clearHyperlink, type SetHyperlinkOptions, setHyperlink } from './hyperlinks.js';
import type { InteractionOperation, InteractionOrigin } from './interaction-policy.js';
import { recordFormatChange } from './slice-history.js';

/** Where a hyperlink change runs and which entry point asked for it. */
export interface HyperlinkActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly workbook?: WorkbookHandle;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The operation of a hyperlink change. Adding or editing a link is `hyperlink`; removing one
 *  clears content, so it is `clear`: on a protected sheet a removal needs only an unlocked cell,
 *  while a new link also needs the Insert hyperlinks option. */
export const hyperlinkOperation = (removes: boolean): InteractionOperation =>
  removes ? 'clear' : 'hyperlink';

/**
 * Run a hyperlink change of `cells` as one authorized step; `removes` marks a removal. The link,
 * its display text and its ScreenTip live in the cells' format, not their values, so the change
 * is one intent on the cells. `write` records its own undo entry.
 */
export function hyperlinkWrite<T>(
  ctx: HyperlinkActionContext,
  cells: readonly Addr[],
  refused: T,
  write: () => T,
  removes = false,
): T {
  if (cells.length === 0) return refused;
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents: (origin) => [
      {
        operation: hyperlinkOperation(removes),
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

/** Set the hyperlink of `addr` (or remove it when `target` is empty), as one undoable step. */
export function setHyperlinkAction(
  ctx: HyperlinkActionContext,
  addr: Addr,
  target: string | undefined,
  options: SetHyperlinkOptions = {},
): boolean {
  return hyperlinkWrite(
    ctx,
    [addr],
    false,
    () => {
      recordFormatChange(ctx.history, ctx.store, () => {
        if (target) setHyperlink(ctx.store, addr, target, ctx.workbook, options);
        else clearHyperlink(ctx.store, addr, ctx.workbook);
      });
      return true;
    },
    !target,
  );
}
