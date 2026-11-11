import type { Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../../store/store.js';
import { type AuthorizedIntentFactory, authorizedWrite } from '../authorized-write.js';
import { History } from '../history.js';
import type { InteractionOrigin } from '../interaction-policy.js';
import {
  copiedBlock,
  type InsertCopiedBandTarget,
  type InsertCopiedCellsDirection,
  type InsertCopiedCellsResult,
  insertCopiedBand,
  insertCopiedCellsFromTSV,
  targetAddress,
} from './insert-copied-cells.js';
import { resolvePasteDestination } from './paste-destination.js';
import {
  insertCopiedBandIntents,
  insertCopiedBlockIntents,
  pasteIntents,
} from './paste-intents.js';
import { type PasteSpecialOptions, pasteSpecial } from './paste-special.js';
import type { ClipboardSnapshot } from './snapshot.js';

/** Where a paste runs and which entry point asked for it. */
export interface PasteActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** Run `write` as one authorized step of `intents`, whose undo and redo replay under the same
 *  check; `refused` when there is nothing to authorize or the step is refused. */
function runPaste<T>(
  ctx: PasteActionContext,
  intents: AuthorizedIntentFactory,
  refused: T,
  write: () => T,
): T {
  if (intents(ctx.origin).length === 0) return refused;
  const result = authorizedWrite({
    store: ctx.store,
    history: ctx.history ?? new History(),
    intents,
    origin: ctx.origin,
    write,
  });
  return result.ok ? result.value : refused;
}

/**
 * Run a paste as one authorized step (see {@link pasteIntents}): the whole destination and, for a
 * cut, the source it clears, plus the tables it grows. `extra` adds the intents of what the paste
 * does beyond cells (column widths). `write` records its own undo entries; `refused` is returned
 * when the paste is not allowed.
 */
export function pasteWrite<T>(
  ctx: PasteActionContext,
  ranges: { readonly destination: Range | null; readonly cutSource?: Range | null },
  refused: T,
  write: () => T,
  extra: AuthorizedIntentFactory = () => [],
): T {
  const state = ctx.store.getState();
  return runPaste(
    ctx,
    (origin) => [
      ...pasteIntents(state, origin, ctx.commandId, ranges.destination, ranges.cutSource ?? null),
      ...extra(origin),
    ],
    refused,
    write,
  );
}

/** Paste Special of `snap` from the entry point `ctx` names, as one authorized step. Column
 *  widths write no cell and are decided as a column layout change of the destination's sheet. */
export function pasteSpecialAction(
  ctx: PasteActionContext,
  wb: WorkbookHandle,
  snap: ClipboardSnapshot,
  opts: PasteSpecialOptions,
): ReturnType<typeof pasteSpecial> {
  const state = ctx.store.getState();
  const destination = resolvePasteDestination(state, snap, opts.transpose, opts.what === 'formats');
  if (opts.what === 'column-widths') {
    const sheet = destination?.sheet ?? state.data.sheetIndex;
    return pasteWrite(
      ctx,
      { destination: null },
      null,
      () => pasteSpecial(ctx.store.getState(), ctx.store, wb, snap, opts, ctx.history),
      (origin) => [
        {
          operation: 'resizeColumns',
          origin,
          commandId: ctx.commandId,
          effects: [{ kind: 'sheet', sheet }],
        },
      ],
    );
  }
  return pasteWrite(
    ctx,
    { destination, cutSource: snap.mode === 'cut' ? snap.range : null },
    null,
    () => pasteSpecial(ctx.store.getState(), ctx.store, wb, snap, opts, ctx.history),
  );
}

/** Insert Copied (Cut) Cells of a copied whole-row or whole-column band at `target`, as one
 *  authorized step of {@link insertCopiedBandIntents}. */
export function insertCopiedBandAction(
  ctx: PasteActionContext,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot | null | undefined,
  target?: InsertCopiedBandTarget,
): InsertCopiedCellsResult | null {
  if (!snapshot) return null;
  const state = ctx.store.getState();
  const at = targetAddress(state, target);
  if (!at) return null;
  const source = snapshot.logicalRange ?? snapshot.range;
  const cut = snapshot.mode === 'cut';
  if (insertCopiedBandIntents(state, ctx.origin, ctx.commandId, source, cut, at) === null)
    return null;
  return runPaste(
    ctx,
    (origin) => insertCopiedBandIntents(state, origin, ctx.commandId, source, cut, at) ?? [],
    null,
    () => insertCopiedBand(ctx.store, wb, ctx.history, snapshot, target),
  );
}

/** Insert Copied (Cut) Cells of a block at the active cell, shifting cells `direction`, as one
 *  authorized step of {@link insertCopiedBlockIntents}. */
export function insertCopiedCellsAction(
  ctx: PasteActionContext,
  wb: WorkbookHandle,
  text: string,
  direction: InsertCopiedCellsDirection,
  snapshot?: ClipboardSnapshot | null,
): InsertCopiedCellsResult | null {
  const block = copiedBlock(text, snapshot);
  if (!block) return null;
  const state = ctx.store.getState();
  const size = { rows: block.height, cols: block.width };
  const cutSource = snapshot?.mode === 'cut' ? snapshot.range : null;
  const at = state.selection.active;
  return runPaste(
    ctx,
    (origin) =>
      insertCopiedBlockIntents(state, origin, ctx.commandId, size, cutSource, at, direction),
    null,
    () => insertCopiedCellsFromTSV(ctx.store, wb, ctx.history, text, direction, snapshot),
  );
}
