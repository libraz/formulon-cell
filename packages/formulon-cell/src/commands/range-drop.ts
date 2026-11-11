import { formatA1Cell, MAX_COL, MAX_ROW, quoteSheetName } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangesIntersect } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { copy } from './clipboard/copy.js';
import { createCopyFormatResolver } from './clipboard/format-resolver.js';
import {
  insertCopiedCellsAction,
  pasteSpecialAction,
  pasteWrite,
} from './clipboard/paste-actions.js';
import { resolvePasteDestination } from './clipboard/paste-destination.js';
import { type PasteWhat, pasteLink } from './clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from './clipboard/snapshot.js';
import { inputLocaleForWorkbook, localizeFormulaForClipboard } from './coerce-input.js';
import { History } from './history.js';
import { setHyperlinkAction } from './hyperlink-actions.js';
import { moveSelectionTo } from './move-cells.js';
import { executeRibbonClearAction } from './ribbon-clear.js';

/** What dropping a dragged selection with the secondary button can do at the target. */
export type RangeDropAction =
  | 'move'
  | 'copy'
  | 'copyValues'
  | 'copyFormats'
  | 'link'
  | 'hyperlink'
  | 'copyShiftDown'
  | 'copyShiftRight'
  | 'moveShiftDown'
  | 'moveShiftRight';

export interface RangeDropDeps {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  history: History | null;
}

const PASTE_KIND: Partial<Record<RangeDropAction, PasteWhat>> = {
  copy: 'all',
  copyValues: 'values',
  copyFormats: 'formats',
};

/**
 * Apply a secondary-button drop of `src` onto `dest` (same size, `dest`'s top-left
 * is the drop point). Every action is one undo step and leaves `dest` selected
 * with its first cell active; the drop never asks to confirm overwriting data.
 * Resolves false when nothing changed.
 */
export async function applyRangeDrop(
  deps: RangeDropDeps,
  src: Range,
  dest: Range,
  action: RangeDropAction,
): Promise<boolean> {
  const { store, wb, history } = deps;
  mutators.setActive(store, { sheet: src.sheet, row: src.r0, col: src.c0 });
  mutators.setRange(store, src);

  if (action === 'move') {
    return moveSelectionTo(
      { store, wb, history, commandId: action, confirmOverwrite: async () => true },
      { row: dest.r0, col: dest.c0 },
    );
  }

  const state = store.getState();
  const copied = copy(state, inputLocaleForWorkbook(wb), localizeFormulaForClipboard(wb));
  if (!copied) return false;
  const snap = captureSnapshotFromCopyResult(
    state,
    copied,
    'copy',
    createCopyFormatResolver(wb, state),
  );
  if (!snap) return false;
  const first = { sheet: dest.sheet, row: dest.r0, col: dest.c0 };

  if (action === 'hyperlink') {
    const target = `#${quoteSheetName(wb.sheetName(src.sheet))}!${formatA1Cell(src.r0, src.c0)}`;
    mutators.setActive(store, first);
    mutators.setRange(store, dest);
    return setHyperlinkAction(
      {
        store,
        history: history ?? null,
        workbook: wb,
        origin: 'editor',
        commandId: 'insertHyperlink',
      },
      first,
      target,
    );
  }

  const shift =
    action === 'copyShiftDown' || action === 'moveShiftDown'
      ? 'down'
      : action === 'copyShiftRight' || action === 'moveShiftRight'
        ? 'right'
        : null;
  const movesSource = action === 'moveShiftDown' || action === 'moveShiftRight';
  if (shift && movesSource) {
    // The cells that shift must not include the source it is cleared from.
    const shifted: Range = shift === 'down' ? { ...dest, r1: MAX_ROW } : { ...dest, c1: MAX_COL };
    if (rangesIntersect(shifted, src)) return false;
  }

  mutators.setActive(store, first);
  // Each drop is one authorized step of the gesture, named by the dropped action.
  const ctx = { store, history, origin: 'editor', commandId: action } as const;
  let written: Range | null = null;
  if (shift && !movesSource) {
    written = insertCopiedCellsAction(ctx, wb, copied.tsv, shift, snap)?.writtenRange ?? null;
  } else if (shift) {
    // A shifting move inserts the copy, then clears the source, as one step.
    written = pasteWrite(ctx, { destination: null, cutSource: src }, null, () => {
      const inserted =
        insertCopiedCellsAction(ctx, wb, copied.tsv, shift, snap)?.writtenRange ?? null;
      if (inserted) {
        mutators.setActive(store, { sheet: src.sheet, row: src.r0, col: src.c0 });
        mutators.setRange(store, src);
        executeRibbonClearAction({
          store,
          workbook: wb,
          history: history ?? new History(),
          action: 'all',
          entry: { origin: ctx.origin, commandId: ctx.commandId },
        });
      }
      return inserted;
    });
  } else if (action === 'link') {
    written = pasteWrite(
      ctx,
      { destination: resolvePasteDestination(store.getState(), snap, false) },
      null,
      () => pasteLink(store.getState(), store, wb, snap, history)?.writtenRange ?? null,
    );
  } else {
    const what = PASTE_KIND[action] ?? 'all';
    written =
      pasteSpecialAction(ctx, wb, snap, {
        what,
        operation: 'none',
        skipBlanks: false,
        transpose: false,
      })?.writtenRange ?? null;
  }
  if (!written) {
    mutators.setActive(store, { sheet: src.sheet, row: src.r0, col: src.c0 });
    mutators.setRange(store, src);
    return false;
  }
  mutators.replaceCells(store, wb.cells(dest.sheet));
  mutators.setActive(store, first);
  mutators.setRange(store, dest);
  return true;
}
