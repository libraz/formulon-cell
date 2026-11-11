import { MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Range } from '../../engine/types.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { isWholeColumnRange, isWholeRowRange } from '../../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../../store/store.js';
import { type AuthorizedIntentFactory, authorizedWrite } from '../authorized-write.js';
import { History } from '../history.js';
import { cellShiftBand, columnShiftBand, rowShiftBand } from '../interaction-effects.js';
import type { InteractionOrigin, OperationEffect } from '../interaction-policy.js';
import { isCellWritable, isSheetProtected } from '../protection.js';
import { makeBand } from './copied-band.js';
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
import { type PasteSpecialOptions, pasteSpecial } from './paste-special.js';
import type { ClipboardSnapshot } from './snapshot.js';

/** Where a paste runs and which entry point asked for it. */
export interface PasteActionContext {
  readonly store: SpreadsheetStore;
  readonly history: History | null;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** Cells a protected sheet's paste may still enumerate one by one. */
const MAX_ENUMERATED_PASTE_CELLS = 100_000;

/**
 * The cells of `range` a paste writes. A snapshot paste is atomic and writes all of them; a text
 * paste (`skipsLocked`) skips the locked cells of a protected sheet, so only its writable cells
 * are declared. Null when it writes none.
 */
export const pastedCells = (
  state: State,
  range: Range,
  skipsLocked: boolean,
): OperationEffect | null => {
  if (!skipsLocked || !isSheetProtected(state, range.sheet)) return { kind: 'range', range };
  const area = (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);
  if (area > MAX_ENUMERATED_PASTE_CELLS) return { kind: 'range', range };
  const cells = [];
  for (let row = range.r0; row <= range.r1; row += 1) {
    for (let col = range.c0; col <= range.c1; col += 1) {
      const addr = { sheet: range.sheet, row, col };
      if (isCellWritable(state, addr)) cells.push(addr);
    }
  }
  return cells.length > 0 ? { kind: 'cells', cells } : null;
};

/**
 * Run a paste as one authorized `paste` step whose undo and redo replay under the same check:
 * the destination cells it writes and, for a cut, the source cells it clears. `extra` adds the
 * intents of what the paste does beyond cells (column widths, an insertion). `write` records its
 * own undo entries; `refused` is returned when the paste is not allowed.
 */
export function pasteWrite<T>(
  ctx: PasteActionContext,
  ranges: {
    readonly destination: Range | null;
    readonly cutSource?: Range | null;
    readonly skipsLocked?: boolean;
  },
  refused: T,
  write: () => T,
  extra: AuthorizedIntentFactory = () => [],
): T {
  const state = ctx.store.getState();
  const effects = [ranges.destination, ranges.cutSource ?? null]
    .map((range) => (range ? pastedCells(state, range, ranges.skipsLocked === true) : null))
    .filter((effect): effect is OperationEffect => effect !== null);
  const intents: AuthorizedIntentFactory = (origin) => [
    ...(effects.length > 0
      ? [{ operation: 'paste' as const, origin, commandId: ctx.commandId, effects }]
      : []),
    ...extra(origin),
  ];
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

/**
 * Insert Copied (Cut) Cells of a copied whole-row or whole-column band at `target`: the inserted
 * lines are a row (column) insert, a cut also deletes its source lines, and the band is pasted.
 */
export function insertCopiedBandAction(
  ctx: PasteActionContext,
  wb: WorkbookHandle,
  snapshot: ClipboardSnapshot | null | undefined,
  target?: InsertCopiedBandTarget,
): InsertCopiedCellsResult | null {
  if (!snapshot) return null;
  const logical = snapshot.logicalRange ?? snapshot.range;
  const rows = isWholeRowRange(logical);
  const at = targetAddress(ctx.store.getState(), target);
  if (!at || rows === isWholeColumnRange(logical)) return null;
  const count = rows ? logical.r1 - logical.r0 + 1 : logical.c1 - logical.c0 + 1;
  const insertionAt = rows ? at.row : at.col;
  const cut = snapshot.mode === 'cut';
  const lines = (sheet: number, first: number): OperationEffect => ({
    kind: 'range',
    range: rows ? rowShiftBand(sheet, first) : columnShiftBand(sheet, first),
  });
  return pasteWrite(
    ctx,
    { destination: makeBand(at.sheet, rows, insertionAt, count), cutSource: cut ? logical : null },
    null,
    () => insertCopiedBand(ctx.store, wb, ctx.history, snapshot, target),
    (origin) => [
      {
        operation: rows ? 'insertRows' : 'insertColumns',
        origin,
        commandId: ctx.commandId,
        effects: [lines(at.sheet, insertionAt)],
      },
      ...(cut
        ? [
            {
              operation: rows ? ('deleteRows' as const) : ('deleteColumns' as const),
              origin,
              commandId: ctx.commandId,
              effects: [lines(logical.sheet, rows ? logical.r0 : logical.c0)],
            },
          ]
        : []),
    ],
  );
}

/**
 * Insert Copied (Cut) Cells of a block at the active cell, shifting cells down or right: the
 * shifted cells are a cell insert, a cut also closes its source up or left, and the block is
 * pasted.
 */
export function insertCopiedCellsAction(
  ctx: PasteActionContext,
  wb: WorkbookHandle,
  text: string,
  direction: InsertCopiedCellsDirection,
  snapshot?: ClipboardSnapshot | null,
): InsertCopiedCellsResult | null {
  const block = copiedBlock(text, snapshot);
  if (!block) return null;
  const { active } = ctx.store.getState().selection;
  const destination: Range = {
    sheet: active.sheet,
    r0: active.row,
    c0: active.col,
    r1: Math.min(MAX_ROW, active.row + block.height - 1),
    c1: Math.min(MAX_COL, active.col + block.width - 1),
  };
  const down = direction === 'down';
  const cutSource = snapshot?.mode === 'cut' ? snapshot.range : null;
  return pasteWrite(
    ctx,
    { destination, cutSource },
    null,
    () => insertCopiedCellsFromTSV(ctx.store, wb, ctx.history, text, direction, snapshot),
    (origin) => [
      {
        operation: down ? 'insertRows' : 'insertColumns',
        origin,
        commandId: ctx.commandId,
        effects: [{ kind: 'range', range: cellShiftBand(destination, down) }],
      },
      ...(cutSource
        ? [
            {
              operation: down ? ('deleteRows' as const) : ('deleteColumns' as const),
              origin,
              commandId: ctx.commandId,
              effects: [{ kind: 'range' as const, range: cellShiftBand(cutSource, down) }],
            },
          ]
        : []),
    ],
  );
}
