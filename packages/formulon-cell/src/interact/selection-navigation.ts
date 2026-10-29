import { interactionControllerFor } from '../commands/interaction-controller.js';
import { mergeAt } from '../commands/merge.js';
import { MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import {
  clampNavigationAddr,
  isNavigationAddrAllowed,
  navigationBoundsFor,
  navigationPolicyFor,
} from './navigation-policy.js';
import {
  clearTabReturnOrigin,
  resetTabNavigationSession,
  setTabNavigationMemory,
  tabNavigationMemoryFor,
  withTabNavigationSelectionMutation,
} from './tab-navigation-session.js';
import {
  firstVisibleIndex,
  isCellVisible,
  isVisibleAxisIndex,
  nextVisibleIndex,
  stepVisibleWithMerge,
  visibleLogicalStopCount,
} from './visible-navigation.js';

export type SelectionNavigationDirection = 'down' | 'up' | 'right' | 'left';

const contains = (range: Range, addr: Addr): boolean =>
  range.sheet === addr.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const nextAddress = (
  state: State,
  range: Range,
  addr: Addr,
  direction: SelectionNavigationDirection,
): Addr | null => {
  const vertical = direction === 'down' || direction === 'up';
  const fastAxis = vertical ? 'row' : 'col';
  const slowAxis = vertical ? 'col' : 'row';
  const fastDirection: -1 | 1 = direction === 'down' || direction === 'right' ? 1 : -1;
  const slowDirection = fastDirection;
  const fastMin = vertical ? range.r0 : range.c0;
  const fastMax = vertical ? range.r1 : range.c1;
  const slowMin = vertical ? range.c0 : range.r0;
  const slowMax = vertical ? range.c1 : range.r1;
  const fast = vertical ? addr.row : addr.col;
  const slow = vertical ? addr.col : addr.row;
  const nextFast = nextVisibleIndex(state.layout, fastAxis, fast, fastDirection, fastMin, fastMax);
  if (nextFast !== null) {
    return vertical
      ? { sheet: range.sheet, row: nextFast, col: slow }
      : { sheet: range.sheet, row: slow, col: nextFast };
  }

  let nextSlow = nextVisibleIndex(state.layout, slowAxis, slow, slowDirection, slowMin, slowMax);
  if (nextSlow === null) {
    const start = slowDirection > 0 ? slowMin - 1 : slowMax + 1;
    nextSlow = nextVisibleIndex(state.layout, slowAxis, start, slowDirection, slowMin, slowMax);
  }
  if (nextSlow === null) return null;
  const firstFast = firstVisibleIndex(state.layout, fastAxis, fastMin, fastMax, fastDirection);
  if (firstFast === null) return null;
  return vertical
    ? { sheet: range.sheet, row: firstFast, col: nextSlow }
    : { sheet: range.sheet, row: nextSlow, col: firstFast };
};

const addressAtOrAfter = (
  state: State,
  range: Range,
  addr: Addr,
  direction: SelectionNavigationDirection,
): Addr | null => {
  if (isCellVisible(state.layout, addr.row, addr.col)) return addr;
  const vertical = direction === 'down' || direction === 'up';
  const fastAxis = vertical ? 'row' : 'col';
  const slowAxis = vertical ? 'col' : 'row';
  const fastDirection: -1 | 1 = direction === 'down' || direction === 'right' ? 1 : -1;
  const slowDirection = fastDirection;
  const fastMin = vertical ? range.r0 : range.c0;
  const fastMax = vertical ? range.r1 : range.c1;
  const slowMin = vertical ? range.c0 : range.r0;
  const slowMax = vertical ? range.c1 : range.r1;
  const fast = vertical ? addr.row : addr.col;
  const slow = vertical ? addr.col : addr.row;

  if (isVisibleAxisIndex(state.layout, slowAxis, slow)) {
    const atOrAfterFast = !isVisibleAxisIndex(state.layout, fastAxis, fast)
      ? nextVisibleIndex(
          state.layout,
          fastAxis,
          fast - fastDirection,
          fastDirection,
          fastMin,
          fastMax,
        )
      : fast;
    if (atOrAfterFast !== null) {
      return vertical
        ? { sheet: range.sheet, row: atOrAfterFast, col: slow }
        : { sheet: range.sheet, row: slow, col: atOrAfterFast };
    }
  }

  let nextSlow = nextVisibleIndex(state.layout, slowAxis, slow, slowDirection, slowMin, slowMax);
  if (nextSlow === null) {
    nextSlow = nextVisibleIndex(
      state.layout,
      slowAxis,
      slowDirection > 0 ? slowMin - 1 : slowMax + 1,
      slowDirection,
      slowMin,
      slowMax,
    );
  }
  const firstFast = firstVisibleIndex(state.layout, fastAxis, fastMin, fastMax, fastDirection);
  if (nextSlow === null || firstFast === null) return null;
  return vertical
    ? { sheet: range.sheet, row: firstFast, col: nextSlow }
    : { sheet: range.sheet, row: nextSlow, col: firstFast };
};

const mergeIsContained = (range: Range, merge: Range): boolean =>
  merge.sheet === range.sheet &&
  merge.r0 >= range.r0 &&
  merge.r1 <= range.r1 &&
  merge.c0 >= range.c0 &&
  merge.c1 <= range.c1;

/** Jump over the rest of a merge in the requested logical traversal order. */
const afterMerge = (
  range: Range,
  merge: Range,
  direction: SelectionNavigationDirection,
  candidate: Addr,
): Addr => {
  if (direction === 'down') {
    if (merge.r1 < range.r1) return { sheet: range.sheet, row: merge.r1 + 1, col: candidate.col };
    if (merge.r0 === range.r0 && merge.r1 === range.r1) {
      return {
        sheet: range.sheet,
        row: range.r0,
        col: merge.c1 < range.c1 ? merge.c1 + 1 : range.c0,
      };
    }
    if (candidate.col < range.c1) {
      return { sheet: range.sheet, row: range.r0, col: candidate.col + 1 };
    }
    return { sheet: range.sheet, row: range.r0, col: range.c0 };
  }
  if (direction === 'up') {
    if (merge.r0 > range.r0) return { sheet: range.sheet, row: merge.r0 - 1, col: candidate.col };
    if (candidate.col > range.c0) {
      return { sheet: range.sheet, row: range.r1, col: candidate.col - 1 };
    }
    return { sheet: range.sheet, row: range.r1, col: range.c1 };
  }
  if (direction === 'right') {
    if (merge.c1 < range.c1) return { sheet: range.sheet, row: candidate.row, col: merge.c1 + 1 };
    if (merge.c0 === range.c0 && merge.c1 === range.c1) {
      return {
        sheet: range.sheet,
        row: merge.r1 < range.r1 ? merge.r1 + 1 : range.r0,
        col: range.c0,
      };
    }
    if (candidate.row < range.r1) {
      return { sheet: range.sheet, row: candidate.row + 1, col: range.c0 };
    }
    return { sheet: range.sheet, row: range.r0, col: range.c0 };
  }
  if (merge.c0 > range.c0) return { sheet: range.sheet, row: candidate.row, col: merge.c0 - 1 };
  if (candidate.row > range.r0) {
    return { sheet: range.sheet, row: candidate.row - 1, col: range.c1 };
  }
  return { sheet: range.sheet, row: range.r1, col: range.c1 };
};

/**
 * Return the next logical cell inside the primary selected rectangle.
 * Return/Shift+Return use column-major order; Tab/Shift+Tab use row-major
 * order. Traversal wraps and skips merged-cell bodies without materializing
 * the selected area.
 */
export function nextWithinSelection(
  state: State,
  direction: SelectionNavigationDirection,
): Addr | null {
  const range = state.selection.range;
  if (!contains(range, state.selection.active)) return null;
  if (range.r0 === range.r1 && range.c0 === range.c1) return null;

  const activeMerge = mergeAt(state, state.selection.active);
  const start = activeMerge
    ? { sheet: activeMerge.sheet, row: activeMerge.r0, col: activeMerge.c0 }
    : state.selection.active;
  if (!contains(range, start)) return null;

  // A selection occupied by exactly one merge has only one logical stop.
  if (
    activeMerge &&
    activeMerge.r0 === range.r0 &&
    activeMerge.r1 === range.r1 &&
    activeMerge.c0 === range.c0 &&
    activeMerge.c1 === range.c1
  ) {
    return null;
  }

  const logicalStopCount = visibleLogicalStopCount(
    state.layout,
    range,
    state.merges.byAnchor.values(),
  );
  if (logicalStopCount === 0) return null;
  if (logicalStopCount === 1) {
    return activeMerge || isCellVisible(state.layout, start.row, start.col) ? start : null;
  }

  let candidate = nextAddress(state, range, start, direction);
  if (!candidate) return null;
  // Each jump clears at least one merged cell, so a full-sheet selection never scans its area.
  const maxMergeJumps = state.merges.byCell.size + state.merges.byAnchor.size + 1;
  for (let mergeJumps = 0; mergeJumps <= maxMergeJumps; mergeJumps += 1) {
    if (sameAddr(candidate, start)) {
      return isCellVisible(state.layout, start.row, start.col) ? start : null;
    }
    const merge = mergeAt(state, candidate);
    if (!merge) return candidate;

    const mergeAnchor = { sheet: merge.sheet, row: merge.r0, col: merge.c0 };
    const entersMergeAtLogicalStart =
      sameAddr(candidate, mergeAnchor) ||
      (direction === 'up' && candidate.col === merge.c0) ||
      (direction === 'left' && candidate.row === merge.r0);
    if (
      mergeIsContained(range, merge) &&
      isCellVisible(state.layout, candidate.row, candidate.col) &&
      entersMergeAtLogicalStart
    ) {
      const startMerge = mergeAt(state, start);
      if (
        !startMerge ||
        startMerge.sheet !== merge.sheet ||
        startMerge.r0 !== merge.r0 ||
        startMerge.c0 !== merge.c0
      ) {
        return mergeAnchor;
      }
    }

    const clippedMerge: Range = {
      sheet: range.sheet,
      r0: Math.max(range.r0, merge.r0),
      c0: Math.max(range.c0, merge.c0),
      r1: Math.min(range.r1, merge.r1),
      c1: Math.min(range.c1, merge.c1),
    };
    const after = afterMerge(range, clippedMerge, direction, candidate);
    candidate = addressAtOrAfter(state, range, after, direction);
    if (!candidate) return null;
  }
  throw new Error('nextWithinSelection: merge traversal did not converge');
}

const STEP: Record<SelectionNavigationDirection, readonly [number, number]> = {
  down: [1, 0],
  up: [-1, 0],
  right: [0, 1],
  left: [0, -1],
};

/** Where Return/Tab moves the active cell, and whether the selected area stays intact. */
export interface SelectionAdvance {
  readonly addr: Addr;
  readonly preserveSelection: boolean;
}

/**
 * Resolve one Return/Tab step: inside the selected rectangle when
 * `traverseSelection` is set and the rectangle has another stop, otherwise a
 * merge-aware single step clamped to `bounds` (the sheet limits by default).
 */
export function nextAdvanceTarget(
  state: State,
  direction: SelectionNavigationDirection,
  traverseSelection: boolean,
  bounds?: Range,
): SelectionAdvance {
  if (traverseSelection) {
    const within = nextWithinSelection(state, direction);
    if (within) return { addr: within, preserveSelection: true };
  }
  const [dRow, dCol] = STEP[direction];
  return {
    addr: stepVisibleWithMerge(
      state,
      state.selection.active,
      dRow,
      dCol,
      bounds?.r1 ?? MAX_ROW,
      bounds?.c1 ?? MAX_COL,
      bounds?.r0 ?? 0,
      bounds?.c0 ?? 0,
    ),
    preserveSelection: false,
  };
}

function selectionIsOneLogicalCell(state: State): boolean {
  const selection = state.selection;
  if ((selection.extraRanges?.length ?? 0) !== 0) return false;

  const { active, range } = selection;
  if (range.sheet !== active.sheet) return false;
  const activeKey = `${active.sheet}:${active.row}:${active.col}`;
  const mergeKey = state.merges.byCell.get(activeKey) ?? activeKey;
  const merge = state.merges.byAnchor.get(mergeKey);
  if (merge) {
    return (
      active.row === merge.r0 &&
      active.col === merge.c0 &&
      range.sheet === merge.sheet &&
      range.r0 === merge.r0 &&
      range.r1 === merge.r1 &&
      range.c0 === merge.c0 &&
      range.c1 === merge.c1
    );
  }
  return (
    active.row === range.r0 &&
    active.row === range.r1 &&
    active.col === range.c0 &&
    active.col === range.c1
  );
}

/** Whether the mounted Mac grid can use the transient Tab-entry behavior. */
export function canUseTabNavigationSession(store: SpreadsheetStore, mac: boolean): boolean {
  if (!mac) return false;
  if (interactionControllerFor(store)?.policy !== undefined) return false;
  if (navigationPolicyFor(store)?.options !== undefined) return false;
  return selectionIsOneLogicalCell(store.getState());
}

function advanceInTabNavigationSession(
  store: SpreadsheetStore,
  direction: SelectionNavigationDirection,
): boolean {
  if (!canUseTabNavigationSession(store, true)) return false;

  const state = store.getState();
  const source = state.selection.active;
  const memory = tabNavigationMemoryFor(store);

  if (direction === 'up') {
    // Shift+Return cancels both the Return origin and the horizontal lane.
    resetTabNavigationSession(store);
    return false;
  }

  if (direction === 'down') {
    const target = memory?.returnOrigin
      ? stepVisibleWithMerge(state, memory.returnOrigin, 1, 0, MAX_ROW, MAX_COL)
      : nextAdvanceTarget(state, 'down', false).addr;
    withTabNavigationSelectionMutation(store, () => {
      mutators.setActive(store, target);
    });
    // Return ends the horizontal run, including a run that began with only
    // reverse Tabs and therefore had no Return origin.
    resetTabNavigationSession(store);
    return true;
  }

  const forward = direction === 'right';
  const directionCol = forward ? 1 : -1;
  const laneRow = memory?.laneRow ?? source.row;
  const from = { ...source, row: laneRow };
  const target = stepVisibleWithMerge(state, from, 0, directionCol, MAX_ROW, MAX_COL);
  withTabNavigationSelectionMutation(store, () => {
    mutators.setActive(store, target);
  });
  const moved = !sameAddr(source, store.getState().selection.active);
  if (forward) {
    if (moved) {
      setTabNavigationMemory(store, {
        laneRow,
        returnOrigin: memory?.returnOrigin ?? source,
      });
    }
  } else {
    clearTabReturnOrigin(store);
    setTabNavigationMemory(store, { laneRow });
  }
  return true;
}

/**
 * Move the active cell after a committed edit. Selected-rectangle traversal is
 * a Mac convention and yields to any interaction or navigation policy; plain
 * steps respect merges, navigation bounds, and disabled selection.
 */
export function advanceAfterCommit(
  store: SpreadsheetStore,
  direction: SelectionNavigationDirection,
  mac: boolean,
): void {
  const controller = interactionControllerFor(store);
  if (controller?.policy?.selection === false) return;
  if (mac && advanceInTabNavigationSession(store, direction)) return;
  const traverse =
    mac && controller?.policy === undefined && navigationPolicyFor(store)?.options === undefined;
  const next = nextAdvanceTarget(store.getState(), direction, traverse, navigationBoundsFor(store));
  if (next.preserveSelection) {
    mutators.setActivePreservingSelection(store, next.addr);
    return;
  }
  const clamped = clampNavigationAddr(store, next.addr);
  if (clamped && isNavigationAddrAllowed(store, clamped)) mutators.setActive(store, clamped);
}
