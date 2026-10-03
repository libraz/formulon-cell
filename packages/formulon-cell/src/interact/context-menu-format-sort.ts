import {
  applyValueFilter,
  clearFilter,
  distinctValues,
  filterValueKey,
  inferAutoFilterRange,
  reapplyFilters,
  recordFilterChange,
} from '../commands/filter.js';
import {
  cycleBorders,
  setAlign,
  toggleBold,
  toggleItalic,
  toggleUnderline,
} from '../commands/format.js';
import type { History } from '../commands/history.js';
import { phoneticReading, setPhoneticReading } from '../commands/phonetic.js';
import { inferSortHasHeader, sortRange } from '../commands/sort.js';
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore, State } from '../store/store.js';
import { showPrompt } from '../toolbar/dialogs/prompt.js';
import type { ItemId } from './context-menu-spec.js';

export interface ContextMenuFormatSortContext {
  readonly store: SpreadsheetStore;
  readonly wb: WorkbookHandle;
  readonly history: History | null;
  /** Read at activation so `setStrings` reaches the phonetic prompt. */
  readonly strings: () => Strings;
  /** Forwards to the host's `onAfterCommit`, looked up at call time. */
  readonly afterCommit: () => void;
  /** Records a repeatable, context-menu-originated format change. */
  readonly formatChange: (commandId: string, fn: () => void) => void;
}

/** Clamp a selection's row span to the populated region. A whole-row /
 *  whole-column band selection spans ~1M rows; sorting or filtering that
 *  raw range would iterate (and rewrite) the entire sheet and freeze the
 *  UI, so bound it to the last populated row in the relevant columns. */
const boundRowsToData = (store: SpreadsheetStore, range: Range): Range => {
  if (range.r1 - range.r0 < 50_000) return range;
  let maxRow = range.r0;
  const state = store.getState();
  const visitKey = (key: string): void => {
    const parts = key.split(':');
    if (parts.length !== 3 || Number(parts[0]) !== range.sheet) return;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (row < range.r0 || row > range.r1) return;
    if (col < range.c0 || col > range.c1) return;
    if (row > maxRow) maxRow = row;
  };
  for (const key of state.data.cells.keys()) visitKey(key);
  for (const [key, format] of state.format.formats) {
    if (Object.keys(format).length === 0) continue;
    visitKey(key);
  }
  return { ...range, r1: maxRow };
};

/** Run a character-format, phonetic, filter or sort item. Returns false for
 *  any other item. */
export function runContextMenuFormatSortItem(
  ctx: ContextMenuFormatSortContext,
  id: ItemId,
  state: State,
): boolean {
  const { store, wb, history, formatChange: wrapFmt } = ctx;
  switch (id) {
    case 'bold': {
      wrapFmt('bold', () => toggleBold(store.getState(), store));
      return true;
    }
    case 'italic': {
      wrapFmt('italic', () => toggleItalic(store.getState(), store));
      return true;
    }
    case 'underline': {
      wrapFmt('underline', () => toggleUnderline(store.getState(), store));
      return true;
    }
    case 'alignLeft': {
      wrapFmt('alignLeft', () => setAlign(store.getState(), store, 'left'));
      return true;
    }
    case 'alignCenter': {
      wrapFmt('alignCenter', () => setAlign(store.getState(), store, 'center'));
      return true;
    }
    case 'alignRight': {
      wrapFmt('alignRight', () => setAlign(store.getState(), store, 'right'));
      return true;
    }
    case 'borders': {
      wrapFmt('borders', () => cycleBorders(store.getState(), store));
      return true;
    }
    case 'editPhonetic': {
      const addr = state.selection.active;
      if (!wb.capabilities.phonetic) return true;
      const initial = phoneticReading(state.format.formats.get(addrKey(addr))?.phonetic);
      const strings = ctx.strings();
      void showPrompt({
        title: strings.contextMenu.phoneticDialogTitle,
        label: strings.contextMenu.phoneticDialogLabel,
        initial,
        okLabel: strings.formatDialog.ok,
        cancelLabel: strings.formatDialog.cancel,
      }).then((phonetic) => {
        if (phonetic === null) return;
        if (!setPhoneticReading(store, wb, addr, phonetic, initial)) return;
        ctx.afterCommit();
      });
      return true;
    }
    case 'filterClear': {
      const range = boundRowsToData(store, state.ui.filterRange ?? inferAutoFilterRange(state));
      recordFilterChange(history, store, () => clearFilter(store.getState(), store, range));
      ctx.afterCommit();
      return true;
    }
    case 'filterReapply': {
      recordFilterChange(history, store, () => reapplyFilters(store.getState(), store));
      ctx.afterCommit();
      return true;
    }
    case 'filterByValue': {
      const range = boundRowsToData(store, state.ui.filterRange ?? inferAutoFilterRange(state));
      const byCol = state.selection.active.col;
      const keep = filterValueKey(state.data.cells.get(addrKey(state.selection.active))?.value);
      const hidden = distinctValues(state, range, byCol).filter((k) => k !== keep);
      recordFilterChange(history, store, () =>
        applyValueFilter(store.getState(), store, range, byCol, hidden),
      );
      ctx.afterCommit();
      return true;
    }
    case 'sortAsc':
    case 'sortDesc': {
      const range = boundRowsToData(store, inferAutoFilterRange(state, state.selection.range));
      // A sort rewrites every cell in the range; group them so one undo
      // restores the original order.
      if (history) history.begin();
      try {
        sortRange(
          state,
          store,
          wb,
          range,
          {
            byCol: state.selection.active.col,
            direction: id === 'sortAsc' ? 'asc' : 'desc',
            hasHeader: inferSortHasHeader(state, range),
          },
          history,
        );
      } finally {
        if (history) history.end();
      }
      ctx.afterCommit();
      return true;
    }
    default:
      return false;
  }
}
