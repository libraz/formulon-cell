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
  resetFilterCriteria,
  setAutoFilter,
} from '../commands/filter.js';
import {
  type FilterActionContext,
  filterDefinitionWrite,
  filterUseWrite,
} from '../commands/filter-actions.js';
import { activeFilterRange } from '../commands/filter-ranges.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Range } from '../engine/types.js';

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
  /** The store, history, workbook and entry point the filter change runs under. */
  context: FilterActionContext;
  action: RibbonFilterDataAction;
}

/** Every row of `sheet`: what Clear (with no filter) and Reapply act on. */
const wholeSheet = (sheet: number): Range => ({ sheet, r0: 0, c0: 0, r1: MAX_ROW, c1: MAX_COL });

/** Toggle the auto-filter on the active range as one undoable step (Ctrl+Shift+L,
 *  the Filter button). Turning AutoFilter on or off is decided for the sheet. */
export const toggleAutoFilterFromSelection = (ctx: FilterActionContext): void => {
  const { store, workbook: wb } = ctx;
  const state = store.getState();
  filterDefinitionWrite(ctx, state.data.sheetIndex, undefined, () => {
    const active = activeFilterRange(state);
    if (active) clearFilter(state, store, active, wb);
    else setAutoFilter(store, inferAutoFilterRange(state), wb);
  });
};

/** Dispatch a "Filter" menu item. The helper performs every mutation here so
 *  hosts only handle UI-side outcomes (opening dialogs / dropdowns):
 *  - `toggle` → mutation already ran; `{kind:'toggle'}` is purely a "refresh
 *    your open/close UI state" signal so hosts don't re-toggle.
 *  - `clear` / `reapply` / `filter-by-selected` → uses of the filter.
 *  - `advanced` → return the inferred range so the host can prefill the
 *    advanced-filter dialog.
 *  - default → ensure an auto-filter exists, then ask the host to open the
 *    filter dropdown at the active column. */
export const executeRibbonFilterDataAction = (
  deps: ExecuteRibbonFilterDataActionDeps,
): RibbonFilterDataActionResult => {
  const { context: ctx, action } = deps;
  const { store, workbook: wb } = ctx;
  const state = store.getState();
  if (action === 'toggle') {
    toggleAutoFilterFromSelection(ctx);
    return { kind: 'toggle' };
  }
  if (action === 'clear') {
    const active = activeFilterRange(state);
    // With no filter at all, Clear reveals every row; otherwise it never
    // reaches past the active filter into another table's.
    if (active || state.ui.filterRanges.length === 0) {
      filterUseWrite(ctx, active ?? wholeSheet(state.data.sheetIndex), undefined, () =>
        active ? resetFilterCriteria(store, active, wb) : clearFilter(state, store, undefined, wb),
      );
    }
    return { kind: 'mutated' };
  }
  if (action === 'reapply') {
    filterUseWrite(ctx, wholeSheet(state.data.sheetIndex), undefined, () =>
      reapplyFilters(store.getState(), store, wb),
    );
    return { kind: 'mutated' };
  }
  if (action === 'filter-by-selected') {
    const range = activeFilterRange(state) ?? inferAutoFilterRange(state);
    filterUseWrite(ctx, range, undefined, () =>
      filterBySelectedCellValue(store.getState(), store, undefined, wb),
    );
    return { kind: 'mutated' };
  }
  if (action === 'advanced') {
    return {
      kind: 'open-advanced',
      range: activeFilterRange(state) ?? inferAutoFilterRange(state),
    };
  }
  const active = activeFilterRange(state);
  const range = active ?? inferAutoFilterRange(state);
  if (!active) {
    filterDefinitionWrite(ctx, range.sheet, undefined, () => setAutoFilter(store, range, wb));
  }
  return { kind: 'open-filter-dropdown', range, column: state.selection.active.col };
};
