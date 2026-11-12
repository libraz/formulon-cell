import { cellShiftChoices } from '../commands/built-in-command-policy.js';
import type { History } from '../commands/history.js';
import { groupCols, groupRows, ungroupCols, ungroupRows } from '../commands/outline.js';
import {
  deleteCellsAction,
  deleteColumnsAction,
  deleteRowsAction,
  insertCellsAction,
  insertColumnsAction,
  insertRowsAction,
  insertTableColumnsAction,
  insertTableRowsAction,
  layoutWrite,
  type RowColActionContext,
} from '../commands/row-col-actions.js';
import {
  hiddenExtentInSelection,
  hideCols,
  hideRows,
  showCols,
  showRows,
} from '../commands/row-col-layout.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import {
  isWholeColumnRange,
  isWholeRowRange,
  normalizeRange,
} from '../store/selection-geometry.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { openCellShiftDialog, runCellShiftChoice } from './cell-shift-dialog.js';
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

const wholeLines = (range: Range): boolean => {
  const r = normalizeRange(range);
  return isWholeRowRange(r) || isWholeColumnRange(r);
};

/** Run a cell-shift or row/column structure item against the selection the
 *  menu captured on activation. Returns false for any other item. */
export function runContextMenuStructureItem(
  ctx: ContextMenuStructureContext,
  id: ItemId,
  state: State,
): boolean {
  const { host, store, wb, history } = ctx;
  // Every item runs as itself, from the context menu.
  const action: RowColActionContext = {
    store,
    workbook: wb,
    history,
    origin: 'contextMenu',
    commandId: id,
  };
  const shift = { store, wb, history, origin: 'contextMenu', commandId: id } as const;
  switch (id) {
    case 'insertCells': {
      // Whole rows or columns are inserted at once, without the dialog.
      if (wholeLines(state.selection.range)) {
        if (insertCellsAction(action, state.selection.range, 'down')) ctx.afterCommit();
        return true;
      }
      openCellShiftDialog({
        host,
        strings: ctx.strings(),
        kind: 'insert',
        allowed: cellShiftChoices(store, 'insert', 'contextMenu'),
        onSubmit: (choice) => {
          if (runCellShiftChoice(shift, 'insert', state.selection.range, choice)) ctx.afterCommit();
        },
      });
      return true;
    }
    case 'insertTableRows':
    case 'insertTableColumns': {
      const insert = id === 'insertTableRows' ? insertTableRowsAction : insertTableColumnsAction;
      if (insert(action, state.selection.range)) ctx.afterCommit();
      return true;
    }
    case 'deleteCells': {
      if (wholeLines(state.selection.range)) {
        if (deleteCellsAction(action, state.selection.range, 'up')) ctx.afterCommit();
        return true;
      }
      openCellShiftDialog({
        host,
        strings: ctx.strings(),
        kind: 'delete',
        allowed: cellShiftChoices(store, 'delete', 'contextMenu'),
        onSubmit: (choice) => {
          if (runCellShiftChoice(shift, 'delete', state.selection.range, choice)) ctx.afterCommit();
        },
      });
      return true;
    }
    case 'rowInsertAbove': {
      const r = state.selection.range;
      if (insertRowsAction(action, r.r0, r.r1 - r.r0 + 1)) ctx.afterCommit();
      return true;
    }
    case 'rowInsertBelow': {
      const r = state.selection.range;
      if (insertRowsAction(action, r.r1 + 1, r.r1 - r.r0 + 1)) ctx.afterCommit();
      return true;
    }
    case 'rowDelete': {
      const r = state.selection.range;
      if (deleteRowsAction(action, r.r0, r.r1 - r.r0 + 1)) ctx.afterCommit();
      return true;
    }
    case 'rowHide': {
      const r = state.selection.range;
      layoutWrite(action, 'row', undefined, () => hideRows(store, history, r.r0, r.r1));
      return true;
    }
    case 'rowUnhide': {
      const r = state.selection.range;
      const hidden = hiddenExtentInSelection(state.layout, 'row', r.r0, r.r1);
      if (!hidden) return true;
      layoutWrite(action, 'row', undefined, () =>
        showRows(store, history, hidden.first, hidden.last),
      );
      return true;
    }
    case 'rowGroup': {
      const r = state.selection.range;
      layoutWrite(action, 'row', undefined, () => groupRows(store, history, r.r0, r.r1));
      return true;
    }
    case 'rowUngroup': {
      const r = state.selection.range;
      layoutWrite(action, 'row', undefined, () => ungroupRows(store, history, r.r0, r.r1));
      return true;
    }
    case 'colInsertLeft': {
      const r = state.selection.range;
      if (insertColumnsAction(action, r.c0, r.c1 - r.c0 + 1)) ctx.afterCommit();
      return true;
    }
    case 'colInsertRight': {
      const r = state.selection.range;
      if (insertColumnsAction(action, r.c1 + 1, r.c1 - r.c0 + 1)) ctx.afterCommit();
      return true;
    }
    case 'colDelete': {
      const r = state.selection.range;
      if (deleteColumnsAction(action, r.c0, r.c1 - r.c0 + 1)) ctx.afterCommit();
      return true;
    }
    case 'colHide': {
      const r = state.selection.range;
      layoutWrite(action, 'col', undefined, () => hideCols(store, history, r.c0, r.c1));
      return true;
    }
    case 'colUnhide': {
      const r = state.selection.range;
      const hidden = hiddenExtentInSelection(state.layout, 'col', r.c0, r.c1);
      if (!hidden) return true;
      layoutWrite(action, 'col', undefined, () =>
        showCols(store, history, hidden.first, hidden.last),
      );
      return true;
    }
    case 'colGroup': {
      const r = state.selection.range;
      layoutWrite(action, 'col', undefined, () => groupCols(store, history, r.c0, r.c1));
      return true;
    }
    case 'colUngroup': {
      const r = state.selection.range;
      layoutWrite(action, 'col', undefined, () => ungroupCols(store, history, r.c0, r.c1));
      return true;
    }
    default:
      return false;
  }
}
