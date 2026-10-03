import {
  insertCopiedBand,
  insertCopiedCellsFromTSV,
} from '../commands/clipboard/insert-copied-cells.js';
import { pasteTSV } from '../commands/clipboard/paste.js';
import { pasteSpecial } from '../commands/clipboard/paste-special.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import type { History } from '../commands/history.js';
import { insertCols, insertRows } from '../commands/structure.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import {
  type ContextMenuClipboard,
  consumeCutMarquee,
  hasPastePayload,
  readClipboard,
} from './context-menu-clipboard.js';
import type { MenuKind } from './context-menu-spec.js';
import { openInsertCopiedCellsDialog } from './insert-copied-cells-dialog.js';

export interface ContextMenuInsertCopiedContext {
  readonly host: HTMLElement;
  readonly store: SpreadsheetStore;
  readonly wb: WorkbookHandle;
  readonly history: History | null;
  /** Read at activation so `setStrings` reaches the direction dialog. */
  readonly strings: () => Strings;
  /** Forwards to the host's `onAfterCommit`, looked up at call time. */
  readonly afterCommit: () => void;
  readonly clipboard: Pick<ContextMenuClipboard, 'snapshot'>;
}

export const wholeBandAxisFor = (snapshot: ClipboardSnapshot): 'row' | 'col' | null => {
  const logical = snapshot.logicalRange ?? snapshot.range;
  const wholeRow = logical.c0 === 0 && logical.c1 >= MAX_COL;
  const wholeCol = logical.r0 === 0 && logical.r1 >= MAX_ROW;
  // A full-sheet selection satisfies both predicates but has no single header
  // axis. Leave it on the ordinary path rather than presenting a misleading
  // row/column-specific insert action.
  if (wholeRow === wholeCol) return null;
  if (wholeRow) return 'row';
  if (wholeCol) return 'col';
  return null;
};

/** "Insert Copied Cells" from a row or column header. The header already
 *  fixes the shift direction, so instead of asking which way to push cells
 *  it opens as many whole rows/columns as the copied band is deep/wide and
 *  drops the copy into them. */
function runInsertCopiedBand(ctx: ContextMenuInsertCopiedContext, kind: 'row' | 'col'): void {
  const { store, wb, history } = ctx;
  const source = store.getState().ui.copyRange;
  if (!source) return;

  const sourceAxisMatchesHeader = (snap: ClipboardSnapshot): boolean => {
    return wholeBandAxisFor(snap) === kind;
  };

  const insertBand = (snap: ClipboardSnapshot | null, text: string): void => {
    if (snap && wholeBandAxisFor(snap) !== null) {
      if (!sourceAxisMatchesHeader(snap)) return;
      // The shared path performs the structural edit, dimension copy, and
      // full-band paste as one transaction. A failed preflight must not
      // fall through to the legacy direction-based inserter because that
      // would leave a different shape behind.
      const result = insertCopiedBand(store, wb, history, snap, store.getState().selection.range);
      if (result) {
        if (snap.mode === 'cut') consumeCutMarquee(store);
        mutators.setRange(store, result.writtenRange);
        ctx.afterCommit();
      }
      return;
    }

    // The header band the user right-clicked. It stays selected afterwards:
    // the inserted rows/columns occupy exactly those indices, and the copy
    // marquee stays up so the same source can be inserted again.
    const target = store.getState().selection.range;
    const count = kind === 'col' ? source.c1 - source.c0 + 1 : source.r1 - source.r0 + 1;
    const band: Range =
      kind === 'col'
        ? { ...target, c1: target.c0 + count - 1 }
        : { ...target, r1: target.r0 + count - 1 };
    if (history) history.begin();
    try {
      let inserted = false;
      if (kind === 'col') {
        inserted = insertCols(store, wb, history, target.c0, count);
      } else {
        inserted = insertRows(store, wb, history, target.r0, count);
      }
      if (!inserted) return;
      mutators.setActive(store, {
        sheet: target.sheet,
        row: kind === 'row' ? target.r0 : (snap?.range.r0 ?? 0),
        col: kind === 'col' ? target.c0 : (snap?.range.c0 ?? 0),
      });
      const next = store.getState();
      if (snap) {
        pasteSpecial(
          next,
          store,
          wb,
          snap,
          { what: 'all', operation: 'none', skipBlanks: false, transpose: false },
          history,
        );
      } else {
        pasteTSV(next, wb, text);
      }
    } catch (err) {
      console.warn('formulon-cell: insert copied cells failed', err);
    } finally {
      if (history) history.end();
    }
    mutators.setRange(store, band);
    ctx.afterCommit();
  };

  // The internal snapshot wins over the TSV text so formats ride along, and
  // taking it first keeps the action synchronous — no clipboard-permission
  // round trip for a copy that came from this grid.
  const snap = ctx.clipboard.snapshot();
  if (snap) {
    insertBand(snap, '');
    return;
  }
  void readClipboard().then((text) => {
    if (text.length > 0) insertBand(null, text);
  });
}

/** "Insert Copied Cells": a row/column header inserts whole bands directly;
 *  a cell menu asks which way to shift before inserting. */
export function runContextMenuInsertCopiedCells(
  ctx: ContextMenuInsertCopiedContext,
  menuKind: MenuKind,
): void {
  const { store, wb, history } = ctx;
  if (menuKind !== 'cell') {
    runInsertCopiedBand(ctx, menuKind);
    return;
  }
  openInsertCopiedCellsDialog({
    host: ctx.host,
    strings: ctx.strings(),
    onSubmit: (direction) => {
      void readClipboard().then((text) => {
        const snap = ctx.clipboard.snapshot();
        if (!hasPastePayload(text, snap)) return;
        const r = insertCopiedCellsFromTSV(store, wb, history, text, direction, snap);
        if (r) {
          // Marquee stays up, same as the row/column header variant.
          mutators.setRange(store, r.writtenRange);
          ctx.afterCommit();
        }
      });
    },
  });
}
