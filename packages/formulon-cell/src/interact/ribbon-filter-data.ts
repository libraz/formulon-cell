// Shared "Filter" ribbon split-button. Several branches are pure store mutations
// (clear / reapply / filter-by-selected) so we run them here, but `toggle`,
// `advanced` (open dialog) and the default "open the filter dropdown at the
// active column" need host-side wiring. Hosts decode the discriminated union
// and dispatch their own dialog/dropdown UI.

import {
  clearFilter,
  filterBySelectedCellValue,
  inferAutoFilterRange,
  reapplyFilters,
  recordFilterChange,
  setAutoFilter,
} from '../commands/filter.js';
import type { History } from '../commands/history.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';

export type RibbonFilterDataAction =
  | 'toggle'
  | 'clear'
  | 'reapply'
  | 'filter-by-selected'
  | 'advanced'
  | 'open';

export type RibbonFilterDataActionResult =
  | { kind: 'toggle' }
  | { kind: 'open-advanced'; range: Range }
  | { kind: 'open-filter-dropdown'; range: Range; column: number }
  | { kind: 'mutated' };

export interface ExecuteRibbonFilterDataActionDeps {
  store: SpreadsheetStore;
  history: History;
  action: RibbonFilterDataAction;
  getWb?: () => WorkbookHandle;
}

/** Toggle the auto-filter on the active range as one undoable step. Used as
 *  the host's "Filter" toolbar callback; the [[RibbonFilterDataAction.toggle]]
 *  branch routes back here so the host can also surface a closeDropdown.
 *  Pulled out so both hosts share the exact same begin/end shape. */
export const toggleAutoFilterFromSelection = (
  store: SpreadsheetStore,
  history: History,
  getWb?: () => WorkbookHandle,
): void => {
  const state = store.getState();
  const wb = getWb?.();
  recordFilterChange(
    history,
    store,
    () => {
      if (state.ui.filterRange) clearFilter(state, store, state.ui.filterRange, wb);
      else setAutoFilter(store, inferAutoFilterRange(state), wb);
    },
    wb,
  );
};

/** Dispatch a "Filter" menu item. The helper performs every mutation here so
 *  hosts only handle UI-side outcomes (opening dialogs / dropdowns):
 *  - `toggle` → mutation already ran; `{kind:'toggle'}` is purely a "refresh
 *    your open/close UI state" signal so hosts don't re-toggle.
 *  - `clear` / `reapply` / `filter-by-selected` → pure store mutations.
 *  - `advanced` → return the inferred range so the host can prefill the
 *    advanced-filter dialog.
 *  - default → ensure an auto-filter exists, then ask the host to open the
 *    filter dropdown at the active column. */
export const executeRibbonFilterDataAction = (
  deps: ExecuteRibbonFilterDataActionDeps,
): RibbonFilterDataActionResult => {
  const { store, history, action } = deps;
  const state = store.getState();
  if (action === 'toggle') {
    toggleAutoFilterFromSelection(store, history, deps.getWb);
    return { kind: 'toggle' };
  }
  if (action === 'clear') {
    const wb = deps.getWb?.();
    recordFilterChange(
      history,
      store,
      () => clearFilter(state, store, state.ui.filterRange ?? undefined, wb),
      wb,
    );
    return { kind: 'mutated' };
  }
  if (action === 'reapply') {
    const wb = deps.getWb?.();
    recordFilterChange(history, store, () => reapplyFilters(store.getState(), store, wb), wb);
    return { kind: 'mutated' };
  }
  if (action === 'filter-by-selected') {
    const wb = deps.getWb?.();
    recordFilterChange(
      history,
      store,
      () => filterBySelectedCellValue(store.getState(), store, undefined, wb),
      wb,
    );
    return { kind: 'mutated' };
  }
  if (action === 'advanced') {
    return { kind: 'open-advanced', range: state.ui.filterRange ?? inferAutoFilterRange(state) };
  }
  const range = state.ui.filterRange ?? inferAutoFilterRange(state);
  const wb = deps.getWb?.();
  recordFilterChange(
    history,
    store,
    () => {
      if (!state.ui.filterRange) setAutoFilter(store, range, wb);
    },
    wb,
  );
  return { kind: 'open-filter-dropdown', range, column: state.selection.active.col };
};
