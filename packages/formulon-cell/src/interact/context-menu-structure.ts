import { deleteCells, insertCells } from '../commands/cell-shift.js';
import type { History } from '../commands/history.js';
import { groupCols, groupRows, ungroupCols, ungroupRows } from '../commands/outline.js';
import {
  hiddenInSelection,
  hideCols,
  hideRows,
  showCols,
  showRows,
} from '../commands/row-col-layout.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../commands/structure.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { openCellShiftDialog } from './cell-shift-dialog.js';
import type { ItemId } from './context-menu-spec.js';

export interface ContextMenuStructureContext {
  readonly host: HTMLElement;
  readonly store: SpreadsheetStore;
  readonly wb: WorkbookHandle;
  readonly history: History | null;
  /** Read at activation so `setStrings` reaches the shift dialog. */
  readonly strings: () => Strings;
  /** Forwards to the host's `onAfterCommit`, looked up at call time. */
  readonly afterCommit: () => void;
}

/** Run a cell-shift or row/column structure item against the selection the
 *  menu captured on activation. Returns false for any other item. */
export function runContextMenuStructureItem(
  ctx: ContextMenuStructureContext,
  id: ItemId,
  state: State,
): boolean {
  const { host, store, wb, history } = ctx;
  switch (id) {
    case 'insertCells': {
      openCellShiftDialog({
        host,
        strings: ctx.strings(),
        kind: 'insert',
        onSubmit: (direction) => {
          if (direction !== 'down' && direction !== 'right') return;
          if (insertCells(store, wb, history, state.selection.range, direction)) {
            ctx.afterCommit();
          }
        },
      });
      return true;
    }
    case 'deleteCells': {
      openCellShiftDialog({
        host,
        strings: ctx.strings(),
        kind: 'delete',
        onSubmit: (direction) => {
          if (direction !== 'up' && direction !== 'left') return;
          if (deleteCells(store, wb, history, state.selection.range, direction)) {
            ctx.afterCommit();
          }
        },
      });
      return true;
    }
    case 'rowInsertAbove': {
      const r = state.selection.range;
      insertRows(store, wb, history, r.r0, r.r1 - r.r0 + 1);
      ctx.afterCommit();
      return true;
    }
    case 'rowInsertBelow': {
      const r = state.selection.range;
      insertRows(store, wb, history, r.r1 + 1, r.r1 - r.r0 + 1);
      ctx.afterCommit();
      return true;
    }
    case 'rowDelete': {
      const r = state.selection.range;
      deleteRows(store, wb, history, r.r0, r.r1 - r.r0 + 1);
      ctx.afterCommit();
      return true;
    }
    case 'rowHide': {
      const r = state.selection.range;
      hideRows(store, history, r.r0, r.r1);
      return true;
    }
    case 'rowUnhide': {
      const r = state.selection.range;
      const targets = hiddenInSelection(state.layout, 'row', r.r0, r.r1);
      const first = targets[0];
      const last = targets[targets.length - 1];
      if (first === undefined || last === undefined) return true;
      showRows(store, history, first, last);
      return true;
    }
    case 'rowGroup': {
      const r = state.selection.range;
      groupRows(store, history, r.r0, r.r1);
      return true;
    }
    case 'rowUngroup': {
      const r = state.selection.range;
      ungroupRows(store, history, r.r0, r.r1);
      return true;
    }
    case 'colInsertLeft': {
      const r = state.selection.range;
      insertCols(store, wb, history, r.c0, r.c1 - r.c0 + 1);
      ctx.afterCommit();
      return true;
    }
    case 'colInsertRight': {
      const r = state.selection.range;
      insertCols(store, wb, history, r.c1 + 1, r.c1 - r.c0 + 1);
      ctx.afterCommit();
      return true;
    }
    case 'colDelete': {
      const r = state.selection.range;
      deleteCols(store, wb, history, r.c0, r.c1 - r.c0 + 1);
      ctx.afterCommit();
      return true;
    }
    case 'colHide': {
      const r = state.selection.range;
      hideCols(store, history, r.c0, r.c1);
      return true;
    }
    case 'colUnhide': {
      const r = state.selection.range;
      const targets = hiddenInSelection(state.layout, 'col', r.c0, r.c1);
      const first = targets[0];
      const last = targets[targets.length - 1];
      if (first === undefined || last === undefined) return true;
      showCols(store, history, first, last);
      return true;
    }
    case 'colGroup': {
      const r = state.selection.range;
      groupCols(store, history, r.c0, r.c1);
      return true;
    }
    case 'colUngroup': {
      const r = state.selection.range;
      ungroupCols(store, history, r.c0, r.c1);
      return true;
    }
    default:
      return false;
  }
}
