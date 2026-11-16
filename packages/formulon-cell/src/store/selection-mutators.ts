import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import {
  clampNavigationAddr,
  clampNavigationRange,
  isNavigationAddrAllowed,
  navigationPolicyFor,
  navigationSelectionBoundsFor,
} from '../interact/navigation-policy.js';
import { sameAddr } from './direct-format.js';
import {
  applySelectionRectangle,
  cloneRange,
  isValidRange,
  normalizeRange,
  rangeContainsAddr,
  rangeContainsRange,
  rangeOfAddr,
  rangesIntersect,
  type SelectionGestureMode,
  sameRange,
  selectionRangeList,
} from './selection-geometry.js';
import type { SpreadsheetStore } from './store.js';
import type { State } from './types.js';

/** Navigation is optional. Keep legacy mutator behavior until a policy is
 * registered for the store. */
const permittedAddr = (store: SpreadsheetStore, addr: Addr): Addr | null =>
  navigationPolicyFor(store) ? clampNavigationAddr(store, addr) : addr;

const validSelectionRange = (range: Range): boolean =>
  Number.isSafeInteger(range.sheet) && isValidRange(range);

/** Inverted ranges are stored normalized; off-sheet ranges are rejected before
 *  any navigation clamp. */
const permittedRange = (store: SpreadsheetStore, requested: Range): Range | null => {
  const range = normalizeRange(requested);
  if (!validSelectionRange(range)) return null;
  return navigationPolicyFor(store) ? clampNavigationRange(store, range) : range;
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
      range: rangeOfAddr(addr),
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

/** `value` clamped into `lo..hi`, or `lo` when absent. */
const clampInto = (value: number | undefined, lo: number, hi: number): number =>
  Math.max(lo, Math.min(hi, value ?? lo));

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
  const inSheet = bound?.sheet === sheet ? bound : null;
  const axisMin = axis === 'row' ? (inSheet ? inSheet.r0 : 0) : inSheet ? inSheet.c0 : 0;
  const axisMax =
    axis === 'row' ? (inSheet ? inSheet.r1 : MAX_ROW) : inSheet ? inSheet.c1 : MAX_COL;
  const boundedFirst = Math.max(axisMin, Math.min(axisMax, first));
  const boundedLast = Math.max(axisMin, Math.min(axisMax, last));
  const range: Range =
    axis === 'row'
      ? {
          sheet,
          r0: Math.min(boundedFirst, boundedLast),
          c0: inSheet ? inSheet.c0 : 0,
          r1: Math.max(boundedFirst, boundedLast),
          c1: inSheet ? inSheet.c1 : MAX_COL,
        }
      : {
          sheet,
          r0: inSheet ? inSheet.r0 : 0,
          c0: Math.min(boundedFirst, boundedLast),
          r1: inSheet ? inSheet.r1 : MAX_ROW,
          c1: Math.max(boundedFirst, boundedLast),
        };
  return permittedRange(store, range);
};

/** The moving end of a shift-extension: the corner of the primary range
 *  opposite the active cell. The active cell stays at the origin of the
 *  gesture, so the extend end is derived from the range instead of stored. */
export const selectionExtendEnd = (selection: State['selection']): Addr => {
  const { active, range } = selection;
  return {
    sheet: range.sheet,
    row: active.row === range.r0 ? range.r1 : range.r0,
    col: active.col === range.c0 ? range.c1 : range.c0,
  };
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
      if (sameAddr(s.selection.active, next)) return s;
      const prevPrimary = s.selection.range;
      // Demote the current primary range into extraRanges, promote the new
      // cell to primary so future shift-extends widen the new band.
      return {
        ...s,
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
      if (sameRange(prevPrimary, nextRange)) return s;
      return {
        ...s,
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
      const next = applySelectionRectangle(base, gesture, mode, anchor);
      if (!next) return s;

      const active = selectionForAddr(s, next.active).active;
      const anchorAddr = selectionForAddr(s, next.anchor).active;
      const ranges = selectionRangeList(next);
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
        selection: {
          active,
          anchor: anchorAddr,
          range: { ...next.range },
          extraRanges: (next.extraRanges ?? []).map(cloneRange),
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
        selection: {
          ...s.selection,
          active: a,
          anchor: a,
          range: bounded,
        },
      };
    });
  },

  /** Select whole row `row`. `activeCol` places the active cell (clamped into the range); it
   *  defaults to the range's first column. */
  selectRow(store: SpreadsheetStore, row: number, activeCol?: number): void {
    const range = boundedAxisRange(store, 'row', row, row);
    if (!range) return;
    const col = clampInto(activeCol, range.c0, range.c1);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: range.sheet, row: range.r0, col },
        anchor: { sheet: range.sheet, row: range.r0, col },
        range,
        extraRanges: [],
      },
    }));
  },

  selectRows(store: SpreadsheetStore, anchorRow: number, endRow: number, activeCol?: number): void {
    const r0 = Math.min(anchorRow, endRow);
    const r1 = Math.max(anchorRow, endRow);
    const range = boundedAxisRange(store, 'row', r0, r1);
    if (!range) return;
    const col = clampInto(activeCol, range.c0, range.c1);
    const at = {
      sheet: range.sheet,
      row: Math.max(range.r0, Math.min(range.r1, anchorRow)),
      col,
    };
    store.setState((s) => ({
      ...s,
      selection: {
        active: at,
        anchor: { ...at },
        range,
        extraRanges: [],
      },
    }));
  },

  /** Select whole column `col`. `activeRow` places the active cell (clamped into the range); it
   *  defaults to the range's first row. */
  selectCol(store: SpreadsheetStore, col: number, activeRow?: number): void {
    const range = boundedAxisRange(store, 'col', col, col);
    if (!range) return;
    const row = clampInto(activeRow, range.r0, range.r1);
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: range.sheet, row, col: range.c0 },
        anchor: { sheet: range.sheet, row, col: range.c0 },
        range,
        extraRanges: [],
      },
    }));
  },

  selectCols(store: SpreadsheetStore, anchorCol: number, endCol: number, activeRow?: number): void {
    const c0 = Math.min(anchorCol, endCol);
    const c1 = Math.max(anchorCol, endCol);
    const range = boundedAxisRange(store, 'col', c0, c1);
    if (!range) return;
    const row = clampInto(activeRow, range.r0, range.r1);
    const at = {
      sheet: range.sheet,
      row,
      col: Math.max(range.c0, Math.min(range.c1, anchorCol)),
    };
    store.setState((s) => ({
      ...s,
      selection: {
        active: at,
        anchor: { ...at },
        range,
        extraRanges: [],
      },
    }));
  },

  /** Select the whole sheet. `active` places the active cell (clamped into
   *  the range); it defaults to the range's top-left cell. */
  selectAll(store: SpreadsheetStore, active?: { row: number; col: number }): void {
    const requested =
      navigationSelectionBoundsFor(store) ?? fullSheetRange(store.getState().data.sheetIndex);
    const range = permittedRange(store, requested);
    if (!range) return;
    const row = Math.max(range.r0, Math.min(range.r1, active?.row ?? range.r0));
    const col = Math.max(range.c0, Math.min(range.c1, active?.col ?? range.c0));
    store.setState((s) => ({
      ...s,
      selection: {
        active: { sheet: range.sheet, row, col },
        anchor: { sheet: range.sheet, row, col },
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
      selection: { ...s.selection, range: { ...next } },
    }));
  },
};
