import {
  insertCopiedBand,
  insertCopiedCellsFromTSV,
} from '../commands/clipboard/insert-copied-cells.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import type { History } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { readClipboard } from '../interact/context-menu-clipboard.js';
import { openInsertCopiedCellsDialog } from '../interact/insert-copied-cells-dialog.js';
import { isWholeColumnRange, isWholeRowRange } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';

export interface InsertCopiedCellsDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history: History;
  getWb: () => WorkbookHandle;
  getStrings: () => Strings;
  /** True while an interaction policy blocks the insert. */
  isRestricted: () => boolean;
  getClipboardSnapshot: () => ClipboardSnapshot | null;
  refreshCells: () => void;
  updateChrome: () => void;
  invalidate: () => void;
}

export function createInsertCopiedCellsOpener(deps: InsertCopiedCellsDeps): () => void {
  const { host, store, history, getWb, getStrings } = deps;
  return () => {
    if (deps.isRestricted()) return;
    const wb = getWb();

    // Whole-row/whole-column internal copies and cuts use the structural insert
    // command directly. This preserves source formats, merges, formulas,
    // and row/column dimensions; routing them through the TSV dialog
    // would reduce the payload to values and lose the band topology.
    const copied = deps.getClipboardSnapshot();
    const copiedLogical = copied?.logicalRange ?? copied?.range;
    const copiedWholeBand =
      copiedLogical !== undefined &&
      (isWholeRowRange(copiedLogical) || isWholeColumnRange(copiedLogical));
    if (copiedWholeBand && copied) {
      const target = store.getState().selection.range;
      const result = insertCopiedBand(store, wb, history, copied, target);
      if (result) {
        mutators.replaceCells(store, wb.cells(store.getState().data.sheetIndex));
        mutators.setRange(store, result.writtenRange);
        deps.refreshCells();
        deps.updateChrome();
        deps.invalidate();
      }
      // A valid whole-band snapshot must not fall back to the direction
      // dialog when preflight rejects it: Excel leaves the sheet intact.
      return;
    }

    openInsertCopiedCellsDialog({
      host,
      strings: getStrings(),
      onSubmit: (direction) => {
        void readClipboard().then((text) => {
          const snap = deps.getClipboardSnapshot();
          if (!text && !snap) return;
          // The workbook may have been swapped while the clipboard read was pending.
          const result = insertCopiedCellsFromTSV(store, getWb(), history, text, direction, snap);
          if (!result) return;
          // Marquee stays up, same as the context-menu variants.
          mutators.setRange(store, result.writtenRange);
          deps.refreshCells();
          deps.updateChrome();
          deps.invalidate();
        });
      },
    });
  };
}
