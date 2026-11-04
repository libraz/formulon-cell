import { formatA1Cell, MAX_COL, MAX_ROW, quoteSheetName } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangesIntersect } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { copy } from './clipboard/copy.js';
import { createCopyFormatResolver } from './clipboard/format-resolver.js';
import { insertCopiedCellsFromTSV } from './clipboard/insert-copied-cells.js';
import { type PasteWhat, pasteLink, pasteSpecial } from './clipboard/paste-special.js';
import { captureSnapshotFromCopyResult } from './clipboard/snapshot.js';
import { inputLocaleForWorkbook, localizeFormulaForClipboard } from './coerce-input.js';
import { History } from './history.js';
import { setHyperlink } from './hyperlinks.js';
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
      { store, wb, history, confirmOverwrite: async () => true },
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
    history?.begin();
    try {
      setHyperlink(store, first, target, wb);
    } finally {
      history?.end();
    }
    return true;
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
  history?.begin();
  let written: Range | null = null;
  try {
    if (shift) {
      written =
        insertCopiedCellsFromTSV(store, wb, history, copied.tsv, shift, snap)?.writtenRange ?? null;
      if (written && movesSource) {
        mutators.setActive(store, { sheet: src.sheet, row: src.r0, col: src.c0 });
        mutators.setRange(store, src);
        executeRibbonClearAction({
          store,
          workbook: wb,
          history: history ?? new History(),
          action: 'all',
        });
      }
    } else if (action === 'link') {
      written = pasteLink(store.getState(), store, wb, snap, history)?.writtenRange ?? null;
    } else {
      const what = PASTE_KIND[action] ?? 'all';
      written =
        pasteSpecial(
          store.getState(),
          store,
          wb,
          snap,
          { what, operation: 'none', skipBlanks: false, transpose: false },
          history,
        )?.writtenRange ?? null;
    }
  } finally {
    history?.end();
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
