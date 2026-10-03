import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import {
  clampNavigationAddr,
  clampNavigationRange,
  isNavigationAddrAllowed,
  navigationPolicyFor,
  navigationSelectionBoundsFor,
} from '../interact/navigation-policy.js';
import { sameAddr } from './pending-format.js';
import {
  applySelectionRectangle,
  rangeContainsAddr,
  rangeContainsRange,
  rangesIntersect,
  type SelectionGestureMode,
  sameRange,
} from './selection-geometry.js';
import type { SpreadsheetStore } from './store.js';
import type { State } from './types.js';

const clearPendingFormatOnMove = (s: State, nextActive: Addr): State['ui'] =>
  s.ui.pendingFormat && !sameAddr(s.ui.pendingFormat.addr, nextActive)
    ? { ...s.ui, pendingFormat: null }
    : s.ui;

/** Navigation is optional. Keep legacy mutator behavior until a policy is
 * registered for the store. */
const permittedAddr = (store: SpreadsheetStore, addr: Addr): Addr | null =>
  navigationPolicyFor(store) ? clampNavigationAddr(store, addr) : addr;

const validSelectionRange = (range: Range): boolean =>
  Number.isSafeInteger(range.sheet) &&
  range.sheet >= 0 &&
  Number.isSafeInteger(range.r0) &&
  Number.isSafeInteger(range.c0) &&
  Number.isSafeInteger(range.r1) &&
  Number.isSafeInteger(range.c1) &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r1 <= MAX_ROW &&
  range.c1 <= MAX_COL &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1;

/** Off-sheet or inverted ranges are rejected before any navigation clamp. */
const permittedRange = (store: SpreadsheetStore, range: Range): Range | null => {
  if (!validSelectionRange(range)) return null;
  return navigationPolicyFor(store) ? clampNavigationRange(store, range) : { ...range };
};

const mergeRangeAt = (state: State, addr: Addr): Range | null => {
  const key = addrKey(addr);
  const anchorKey = state.merges.byCell.get(key) ?? key;
  return state.merges.byAnchor.get(anchorKey) ?? null;
};

const selectionForAddr = (state: State, addr: Addr): { active: Addr; range: Range } => {
  const merge = mergeRangeAt(state, addr);
  if (!merge) {
    return {
      active: addr,
      range: { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col },
    };
  }
  return {
    active: { sheet: merge.sheet, row: merge.r0, col: merge.c0 },
    range: { ...merge },
  };
};

const hasPartialMerge = (state: State, range: Range): boolean =>
  [...state.merges.byAnchor.values()].some(
    (merge) => rangesIntersect(merge, range) && !rangeContainsRange(range, merge),
  );

const selectionRangeAllowed = (store: SpreadsheetStore, state: State, range: Range): boolean => {
  if (!validSelectionRange(range) || hasPartialMerge(state, range)) return false;
  const accepted = permittedRange(store, range);
  return (
    !!accepted &&
    sameRange(accepted, range) &&
    isNavigationAddrAllowed(store, { sheet: range.sheet, row: range.r0, col: range.c0 }) &&
    isNavigationAddrAllowed(store, { sheet: range.sheet, row: range.r1, col: range.c1 })
  );
};

const exactSelectionAddr = (store: SpreadsheetStore, addr: Addr): boolean => {
  const accepted = permittedAddr(store, addr);
  return !!accepted && sameAddr(accepted, addr) && isNavigationAddrAllowed(store, addr);
};

const fullSheetRange = (sheet: number): Range => ({
  sheet,
  r0: 0,
  c0: 0,
  r1: MAX_ROW,
  c1: MAX_COL,
});

const boundedAxisRange = (
  store: SpreadsheetStore,
  axis: 'row' | 'col',
  first: number,
  last: number,
): Range | null => {
  const sheet = store.getState().data.sheetIndex;
  if (!navigationPolicyFor(store)) {
    return axis === 'row'
      ? { sheet, r0: first, c0: 0, r1: last, c1: MAX_COL }
      : { sheet, r0: 0, c0: first, r1: MAX_ROW, c1: last };
  }
  const bound = navigationSelectionBoundsFor(store);
  const axisMin =
    axis === 'row'
      ? bound?.sheet === sheet
        ? bound.r0
        : 0
      : bound?.sheet === sheet
        ? bound.c0
        : 0;
  const axisMax =
    axis === 'row'
      ? bound?.sheet === sheet
        ? bound.r1
        : MAX_ROW
      : bound?.sheet === sheet
        ? bound.c1
        : MAX_COL;
  const boundedFirst = Math.max(axisMin, Math.min(axisMax, first));
  const boundedLast = Math.max(axisMin, Math.min(axisMax, last));
  const range: Range =
    axis === 'row'
      ? {
          sheet,
          r0: Math.min(boundedFirst, boundedLast),
          c0: bound?.sheet === sheet ? bound.c0 : 0,
          r1: Math.max(boundedFirst, boundedLast),
          c1: bound?.sheet === sheet ? bound.c1 : MAX_COL,
        }
      : {
          sheet,
          r0: bound?.sheet === sheet ? bound.r0 : 0,
          c0: Math.min(boundedFirst, boundedLast),
          r1: bound?.sheet === sheet ? bound.r1 : MAX_ROW,
          c1: Math.max(boundedFirst, boundedLast),
        };
  return permittedRange(store, range);
};

/** Active cell, anchor, primary range and extra ranges. Every mutation is
 *  clamped by the store's navigation policy and expanded to cover merges. */
export const selectionMutators = {
  setActive(store: SpreadsheetStore, addr: Addr): void {
    const clicked = permittedAddr(store, addr);
    if (!clicked) return;
    store.setState((s) => {
      const resolved = selectionForAddr(s, clicked);
      const active = permittedAddr(store, resolved.active);
      const range = permittedRange(store, resolved.range);
      if (!active || !range) return s;
      return {
        ...s,
        ui: clearPendingFormatOnMove(s, active),
        selection: {
          active,
          anchor: active,
          range,
          extraRanges: [],
        },
      };
    });
  },

  /** Move the active cell while retaining the current selection geometry. Used
   *  for navigation through a selected rectangle. A policy that would clamp
   *  the requested cell elsewhere is rejected so this mutation cannot escape
   *  either the primary selection or the permitted navigation address. */
  setActivePreservingSelection(store: SpreadsheetStore, addr: Addr): void {
    const clicked = permittedAddr(store, addr);
    if (!clicked || !sameAddr(clicked, addr)) return;
    store.setState((s) => {
      const resolved = selectionForAddr(s, clicked);
      const active = permittedAddr(store, resolved.active);
      const range = s.selection.range;
      if (
        !active ||
        !sameAddr(active, resolved.active) ||
        active.sheet !== range.sheet ||
        active.row < range.r0 ||
        active.row > range.r1 ||
        active.col < range.c0 ||
        active.col > range.c1
      ) {
        return s;
      }
      return {
        ...s,
        ui: clearPendingFormatOnMove(s, active),
        selection: { ...s.selection, active },
      };
    });
  },

  /** Append a single-cell range to the current multi-selection. The cell
   *  becomes the new active/anchor so a follow-up shift-click extends from it.
   *  No-op if `addr` is the same sheet/row/col as the current active cell. */
  addExtraCell(store: SpreadsheetStore, addr: Addr): void {
    const clicked = permittedAddr(store, addr);
    if (!clicked) return;
    store.setState((s) => {
      const resolved = selectionForAddr(s, clicked);
      const next = permittedAddr(store, resolved.active);
      const nextRange = permittedRange(store, resolved.range);
      if (!next || !nextRange) return s;
      const sameAsActive =
        s.selection.active.sheet === next.sheet &&
        s.selection.active.row === next.row &&
        s.selection.active.col === next.col;
      if (sameAsActive) return s;
      const prevPrimary = s.selection.range;
      // Demote the current primary range into extraRanges, promote the new
      // cell to primary so future shift-extends widen the new band.
      return {
        ...s,
        ui: clearPendingFormatOnMove(s, next),
        selection: {
          active: next,
          anchor: next,
          range: nextRange,
          extraRanges: [...(s.selection.extraRanges ?? []), prevPrimary],
        },
      };
    });
  },

  /** Add a disjoint range, promoting it to the primary selection and demoting
   *  the previous primary into `extraRanges`. Used by Ctrl/Cmd row/column
   *  header selection. */
  addExtraRange(store: SpreadsheetStore, range: Range, active?: Addr): void {
    const nextRange = permittedRange(store, range);
    if (!nextRange) return;
    const nextActive = permittedAddr(
      store,
      active ?? { sheet: nextRange.sheet, row: nextRange.r0, col: nextRange.c0 },
    );
    if (!nextActive) return;
    store.setState((s) => {
      const prevPrimary = s.selection.range;
      const sameAsPrimary =
        prevPrimary.sheet === nextRange.sheet &&
        prevPrimary.r0 === nextRange.r0 &&
        prevPrimary.c0 === nextRange.c0 &&
        prevPrimary.r1 === nextRange.r1 &&
        prevPrimary.c1 === nextRange.c1;
      if (sameAsPrimary) return s;
      return {
        ...s,
        ui: clearPendingFormatOnMove(s, nextActive),
        selection: {
          active: nextActive,
          anchor: nextActive,
          range: { ...nextRange },
          extraRanges: [...(s.selection.extraRanges ?? []), prevPrimary],
        },
      };
    });
  },

  /** Apply one fixed-mode rectangle gesture against its pointerdown snapshot.
   *  Every range is checked as a unit before the single store publication so
   *  navigation and merge restrictions cannot leave a partial marquee. */
  applySelectionRectangle(
    store: SpreadsheetStore,
    base: State['selection'],
    gesture: Range,
    mode: SelectionGestureMode,
    anchor: Addr,
    tip: Addr,
  ): boolean {
    if (
      !validSelectionRange(gesture) ||
      !rangeContainsAddr(gesture, anchor) ||
      !rangeContainsAddr(gesture, tip) ||
      !exactSelectionAddr(store, anchor) ||
      !exactSelectionAddr(store, tip)
    ) {
      return false;
    }

    let changed = false;
    store.setState((s) => {
      if (!selectionRangeAllowed(store, s, gesture)) return s;
      const next = applySelectionRectangle(base, gesture, mode, anchor, tip);
      if (!next) return s;

      const active = selectionForAddr(s, next.active).active;
      const anchorAddr = selectionForAddr(s, next.anchor).active;
      const ranges = [next.range, ...(next.extraRanges ?? [])];
      if (
        !rangeContainsAddr(next.range, active) ||
        !rangeContainsAddr(next.range, anchorAddr) ||
        !exactSelectionAddr(store, active) ||
        !exactSelectionAddr(store, anchorAddr) ||
        ranges.some((range) => !selectionRangeAllowed(store, s, range))
      ) {
        return s;
      }

      const sameGeometry =
        sameAddr(s.selection.active, active) &&
        sameAddr(s.selection.anchor, anchorAddr) &&
        sameRange(s.selection.range, next.range) &&
        (s.selection.extraRanges ?? []).length === (next.extraRanges ?? []).length &&
        (s.selection.extraRanges ?? []).every((range, index) => {
          const candidate = next.extraRanges?.[index];
          return !!candidate && sameRange(range, candidate);
        });
      if (sameGeometry) return s;

      changed = true;
      return {
        ...s,
        ui: sameAddr(s.selection.active, active) ? s.ui : clearPendingFormatOnMove(s, active),
        selection: {
          active,
          anchor: anchorAddr,
          range: { ...next.range },
          extraRanges: (next.extraRanges ?? []).map((range) => ({ ...range })),
        },
      };
    });
    return changed;
  },

  extendRangeTo(store: SpreadsheetStore, to: Addr): void {
    const next = permittedAddr(store, to);
    if (!next) return;
    store.setState((s) => {
      const a = permittedAddr(store, s.selection.anchor);
      if (!a) return s;
      const requested: Range = {
        sheet: next.sheet,
        r0: Math.min(a.row, next.row),
        c0: Math.min(a.col, next.col),
        r1: Math.max(a.row, next.row),
        c1: Math.max(a.col, next.col),
      };
      const bounded = permittedRange(store, requested);
      if (!bounded) return s;
      return {
        ...s,
        ui: clearPendingFormatOnMove(s, next),
        selection: {
          ...s.selection,
          active: next,
          anchor: a,
          range: bounded,
        },
      };
    });
  },

  /** Set entire row/col selection without an active-cell address change.
   *  Used when the user clicks a row/col header. */
  selectRow(store: SpreadsheetStore, row: number): void {
    const range = boundedAxisRange(store, 'row', row, row);
    if (!range) return;
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, pendingFormat: null },
      selection: {
        active: { sheet: range.sheet, row: range.r0, col: range.c0 },
        anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
        range,
        extraRanges: [],
      },
    }));
  },

  selectRows(store: SpreadsheetStore, anchorRow: number, activeRow: number): void {
    const r0 = Math.min(anchorRow, activeRow);
    const r1 = Math.max(anchorRow, activeRow);
    const range = boundedAxisRange(store, 'row', r0, r1);
    if (!range) return;
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, pendingFormat: null },
      selection: {
        active: {
          sheet: range.sheet,
          row: Math.max(range.r0, Math.min(range.r1, activeRow)),
          col: range.c0,
        },
        anchor: {
          sheet: range.sheet,
          row: Math.max(range.r0, Math.min(range.r1, anchorRow)),
          col: range.c0,
        },
        range,
        extraRanges: [],
      },
    }));
  },

  selectCol(store: SpreadsheetStore, col: number): void {
    const range = boundedAxisRange(store, 'col', col, col);
    if (!range) return;
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, pendingFormat: null },
      selection: {
        active: { sheet: range.sheet, row: range.r0, col: range.c0 },
        anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
        range,
        extraRanges: [],
      },
    }));
  },

  selectCols(store: SpreadsheetStore, anchorCol: number, activeCol: number): void {
    const c0 = Math.min(anchorCol, activeCol);
    const c1 = Math.max(anchorCol, activeCol);
    const range = boundedAxisRange(store, 'col', c0, c1);
    if (!range) return;
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, pendingFormat: null },
      selection: {
        active: {
          sheet: range.sheet,
          row: range.r0,
          col: Math.max(range.c0, Math.min(range.c1, activeCol)),
        },
        anchor: {
          sheet: range.sheet,
          row: range.r0,
          col: Math.max(range.c0, Math.min(range.c1, anchorCol)),
        },
        range,
        extraRanges: [],
      },
    }));
  },

  selectAll(store: SpreadsheetStore): void {
    const requested =
      navigationSelectionBoundsFor(store) ?? fullSheetRange(store.getState().data.sheetIndex);
    const range = permittedRange(store, requested);
    if (!range) return;
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, pendingFormat: null },
      selection: {
        active: { sheet: range.sheet, row: range.r0, col: range.c0 },
        anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
        range,
        extraRanges: [],
      },
    }));
  },

  /** Replace the primary selection range without touching active/anchor. Used
   *  by merge-aware navigation to grow a shift-extend so it covers an entire
   *  merge rectangle. */
  setRange(store: SpreadsheetStore, range: Range): void {
    const next = permittedRange(store, range);
    if (!next) return;
    store.setState((s) => ({
      ...s,
      ui: { ...s.ui, pendingFormat: null },
      selection: { ...s.selection, range: { ...next } },
    }));
  },
};
